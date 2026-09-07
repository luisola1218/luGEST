from __future__ import annotations

import copy
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any


class QualityBackendMixin:
    """Legacy adapter for quality; see BACKEND_GUIDE.md."""

    def quality_summary(self) -> dict[str, Any]:
        data = self.ensure_data()
        ncs = [row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict)]
        docs = [row for row in list(data.get("quality_documents", []) or []) if isinstance(row, dict)]
        health = self.quality_data_health()
        open_ncs = [row for row in ncs if str(row.get("estado", "") or "Aberta").strip().lower() not in {"fechada", "cancelada"}]
        overdue = 0
        today = date.today().isoformat()
        for row in open_ncs:
            due = str(row.get("prazo", "") or "").strip()[:10]
            if due and due < today:
                overdue += 1
        supplier_ncs = [
            row
            for row in open_ncs
            if str(row.get("tipo", "") or "").strip().casefold() == "fornecedor"
            or str(row.get("fornecedor_id", "") or row.get("fornecedor_nome", "") or "").strip()
        ]
        blocked_materials = [
            row
            for row in self._quality_iter_delivery_movements(ensure_ids=False)
            if str(row.get("entity_type", "") or "") == "Material"
            and (
                self._parse_float(row.get("pending_qty", 0), 0) > 0
                or not self._quality_status_is_available(row.get("status", ""))
            )
        ]
        return {
            "open_nc": len(open_ncs),
            "overdue_nc": overdue,
            "supplier_nc": len(supplier_ncs),
            "blocked_materials": len(blocked_materials),
            "documents": len(docs),
            "audit_events": len(list(data.get("audit_log", []) or [])),
            "quality_issues": len(list(health.get("issues", []) or [])),
            "updated_at": str(self.desktop_main.now_iso() or "").strip(),
        }

    def quality_data_health(self) -> dict[str, Any]:
        data = self.ensure_data()
        known: dict[str, set[str]] = {
            "Encomenda": {str(row.get("numero", "") or "").strip() for row in list(data.get("encomendas", []) or []) if isinstance(row, dict)},
            "Material": {str(row.get("id", "") or "").strip() for row in list(data.get("materiais", []) or []) if isinstance(row, dict)},
            "Produto": {str(row.get("codigo", "") or "").strip() for row in list(data.get("produtos", []) or []) if isinstance(row, dict)},
            "Fornecedor": {
                str(row.get("id", "") or row.get("nome", "") or "").strip()
                for row in list(data.get("fornecedores", []) or [])
                if isinstance(row, dict)
            },
            "Cliente": {str(row.get("codigo", "") or row.get("nome", "") or "").strip() for row in list(data.get("clientes", []) or []) if isinstance(row, dict)},
            "Documento": {str(row.get("id", "") or row.get("titulo", "") or "").strip() for row in list(data.get("quality_documents", []) or []) if isinstance(row, dict)},
        }
        reception_keys = self._quality_reception_entity_keys()
        issues: list[dict[str, str]] = []
        for row in list(data.get("quality_nonconformities", []) or []):
            if not isinstance(row, dict):
                continue
            entity_type = str(row.get("entidade_tipo", "") or "").strip()
            entity_id = str(row.get("entidade_id", "") or "").strip()
            if entity_type and entity_type != "Livre" and entity_id:
                if entity_type in {"Material", "Produto"} and (entity_type, entity_id) not in reception_keys and str(row.get("origem", "") or "").strip().casefold() == "receção fornecedor":
                    continue
                known_ids = known.get(entity_type)
                if known_ids is not None and entity_id not in known_ids:
                    issues.append({"tipo": "NC", "id": str(row.get("id", "") or ""), "problema": f"Ligacao inexistente: {entity_type} {entity_id}"})
        for row in list(data.get("quality_documents", []) or []):
            if not isinstance(row, dict):
                continue
            path_txt = str(row.get("caminho", "") or "").strip()
            if path_txt and not Path(path_txt).exists():
                issues.append({"tipo": "Documento", "id": str(row.get("id", "") or ""), "problema": f"Ficheiro nao encontrado: {path_txt}"})
        open_nc_ids = {
            str(row.get("id", "") or "").strip()
            for row in list(data.get("quality_nonconformities", []) or [])
            if isinstance(row, dict) and str(row.get("estado", "") or "Aberta").strip().lower() not in {"fechada", "cancelada"}
        }
        for material in list(data.get("materiais", []) or []):
            if not isinstance(material, dict) or not self._material_quality_is_blocked(material):
                continue
            material_id = str(material.get("id", "") or "").strip()
            if ("Material", material_id) not in reception_keys:
                continue
            status_norm = str(material.get("quality_status", "") or material.get("inspection_status", "") or "").strip().casefold()
            if "inspe" in status_norm and not any(token in status_norm for token in ("bloque", "reclam", "rejeit")):
                continue
            nc_id = str(material.get("quality_nc_id", "") or material.get("supplier_claim_id", "") or "").strip()
            if not nc_id:
                issues.append({"tipo": "Material", "id": str(material.get("id", "") or ""), "problema": "Material bloqueado sem NC/reclamacao ligada."})
            elif nc_id not in open_nc_ids:
                issues.append({"tipo": "Material", "id": str(material.get("id", "") or ""), "problema": f"Material bloqueado com NC inexistente/fechada: {nc_id}"})
        return {"issues": issues, "ok": not issues}

    def quality_link_options(self) -> dict[str, list[dict[str, str]]]:
        data = self.ensure_data()
        options: dict[str, list[dict[str, str]]] = {
            "Livre": [{"id": "", "label": ""}],
            "OPP": [],
            "Encomenda": [],
            "Material": [],
            "Produto": [],
            "Fornecedor": [],
            "Cliente": [],
            "Documento": [],
        }
        for row in list(data.get("plano", []) or []):
            if not isinstance(row, dict):
                continue
            opp = str(row.get("opp", "") or row.get("OPP", "") or "").strip()
            if opp:
                options["OPP"].append({"id": opp, "label": f"{opp} | {str(row.get('encomenda', '') or '').strip()}".strip(" |")})
        for enc in list(data.get("encomendas", []) or []):
            numero = str((enc or {}).get("numero", "") or "").strip()
            if numero:
                options["Encomenda"].append({"id": numero, "label": f"{numero} | {str((enc or {}).get('cliente', '') or '').strip()}"})
        for row in list(data.get("materiais", []) or []):
            if not isinstance(row, dict):
                continue
            material_id = str(row.get("id", "") or "").strip()
            if material_id:
                options["Material"].append(
                    {
                        "id": material_id,
                        "label": " | ".join(
                            part
                            for part in (
                                material_id,
                                str(row.get("material", "") or "").strip(),
                                str(row.get("espessura", "") or "").strip(),
                                str(row.get("formato", "") or "").strip(),
                            )
                            if part
                        ),
                    }
                )
        for row in list(data.get("produtos", []) or []):
            if not isinstance(row, dict):
                continue
            code = str(row.get("codigo", "") or "").strip()
            if code:
                options["Produto"].append({"id": code, "label": f"{code} | {str(row.get('descricao', '') or '').strip()}".strip(" |")})
        for row in list(data.get("fornecedores", []) or []):
            supplier_id = str((row or {}).get("id", "") or "").strip()
            name = str((row or {}).get("nome", "") or "").strip()
            if supplier_id or name:
                options["Fornecedor"].append({"id": supplier_id or name, "label": f"{supplier_id} | {name}".strip(" |")})
        for row in list(data.get("clientes", []) or []):
            code = str((row or {}).get("codigo", "") or "").strip()
            name = str((row or {}).get("nome", "") or "").strip()
            if code or name:
                options["Cliente"].append({"id": code or name, "label": f"{code} | {name}".strip(" |")})
        for row in list(data.get("quality_documents", []) or []):
            doc_id = str((row or {}).get("id", "") or "").strip()
            title = str((row or {}).get("titulo", "") or "").strip()
            if doc_id or title:
                options["Documento"].append({"id": doc_id or title, "label": f"{doc_id} | {title}".strip(" |")})
        for rows in options.values():
            rows.sort(key=lambda item: str(item.get("label", "") or item.get("id", "") or "").lower())
        return options

    def _quality_link_label(self, entity_type: str, entity_id: str) -> str:
        entity_type_txt = str(entity_type or "Livre").strip() or "Livre"
        entity_id_txt = str(entity_id or "").strip()
        if not entity_id_txt:
            return ""
        for row in list(self.quality_link_options().get(entity_type_txt, []) or []):
            if str(row.get("id", "") or "").strip() == entity_id_txt:
                return str(row.get("label", "") or "").strip()
        return entity_id_txt

    def _quality_status_code(self, value: Any) -> str:
        raw = str(value or "").strip().casefold()
        if "devol" in raw:
            return "DEVOLVER_FORNECEDOR"
        if "averig" in raw or "analise" in raw or "análise" in raw:
            return "EM_AVERIGUACAO"
        if "rejeit" in raw:
            return "REJEITADO"
        if "aprov" in raw:
            return "APROVADO"
        return "EM_INSPECAO"

    def _quality_status_is_available(self, value: Any) -> bool:
        return self._quality_status_code(value) == "APROVADO"

    def _quality_quarantine_pending_stock(self, item: dict[str, Any], *, kind: str, max_qty: float | None = None) -> bool:
        if not isinstance(item, dict):
            return False
        status = self._quality_status_code(item.get("quality_status", item.get("inspection_status", "")))
        if status == "APROVADO":
            return False
        qty_key = "qty" if str(kind or "").casefold().startswith("prod") else "quantidade"
        current_qty = self._parse_float(item.get(qty_key, 0), 0)
        pending = self._parse_float(item.get("quality_pending_qty", 0), 0)
        if current_qty <= 0:
            return False
        quarantine_qty = current_qty
        if max_qty is not None:
            quarantine_qty = min(current_qty, max(0.0, self._parse_float(max_qty, 0) - pending))
        if quarantine_qty <= 0:
            return False
        item["quality_pending_qty"] = pending + quarantine_qty
        item["quality_received_qty"] = max(self._parse_float(item.get("quality_received_qty", 0), 0), pending + quarantine_qty)
        item[qty_key] = max(0.0, current_qty - quarantine_qty)
        item["quality_blocked"] = True
        item["logistic_status"] = str(item.get("logistic_status", "") or "RECEBIDO").strip()
        item["atualizado_em"] = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        return True

    def _quality_movement_pending_qty(self, movement: dict[str, Any]) -> float:
        qty = self._parse_float(movement.get("qtd", 0), 0)
        approved = self._parse_float(movement.get("quality_approved_qty", 0), 0)
        rejected = self._parse_float(movement.get("quality_rejected_qty", 0), 0)
        status = self._quality_status_code(movement.get("quality_status", movement.get("inspection_status", "")))
        if approved <= 0 and rejected <= 0:
            if status == "APROVADO":
                approved = qty
            elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                rejected = qty
        return max(0.0, qty - approved - rejected)

    def _quality_delivery_movement_id(self, note_number: str, line_index: int, movement_index: int, entity_type: str, entity_id: str) -> str:
        return "|".join(
            str(part or "").strip()
            for part in (
                note_number,
                f"L{int(line_index or 0)}",
                f"M{int(movement_index or 0)}",
                entity_type,
                entity_id,
            )
        )

    def _quality_reception_entity_keys(self) -> set[tuple[str, str]]:
        return {
            (str(row.get("entity_type", "") or "").strip(), str(row.get("entity_id", "") or "").strip())
            for row in self._quality_iter_delivery_movements(ensure_ids=False)
            if str(row.get("entity_type", "") or "").strip() and str(row.get("entity_id", "") or "").strip()
        }

    def _quality_iter_delivery_movements(self, *, ensure_ids: bool = False) -> list[dict[str, Any]]:
        data = self.ensure_data()
        rows: list[dict[str, Any]] = []
        changed = False
        for note in list(data.get("notas_encomenda", []) or []):
            if not isinstance(note, dict):
                continue
            note_number = str(note.get("numero", "") or "").strip()
            for line_index, line in enumerate(list(note.get("linhas", []) or [])):
                if not isinstance(line, dict):
                    continue
                entity_type = "Material" if self.desktop_main.origem_is_materia(line.get("origem", "")) else "Produto"
                fallback_ref = str(line.get("ref", "") or "").strip()
                for movement_index, movement in enumerate(list(line.get("entregas_linha", []) or [])):
                    if not isinstance(movement, dict):
                        continue
                    entity_id = str(movement.get("stock_ref", "") or fallback_ref).strip()
                    if not entity_id:
                        continue
                    movement_id = str(movement.get("quality_movement_id", "") or "").strip()
                    if not movement_id:
                        movement_id = self._quality_delivery_movement_id(note_number, line_index, movement_index, entity_type, entity_id)
                    status = self._quality_status_code(
                        movement.get("quality_status", "")
                        or movement.get("inspection_status", "")
                        or line.get("quality_status", "")
                        or line.get("inspection_status", "")
                        or "EM_INSPECAO"
                    )
                    qty = self._parse_float(movement.get("qtd", 0), 0)
                    approved = self._parse_float(movement.get("quality_approved_qty", 0), 0)
                    rejected = self._parse_float(movement.get("quality_rejected_qty", 0), 0)
                    if approved <= 0 and rejected <= 0:
                        if status == "APROVADO":
                            approved = qty
                        elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                            rejected = qty
                        else:
                            open_nc = self._quality_find_open_nc(
                                {
                                    "origem": "Receção fornecedor",
                                    "referencia": note_number,
                                    "entidade_tipo": entity_type,
                                    "entidade_id": entity_id,
                                }
                            )
                            if open_nc is not None:
                                approved = min(qty, self._quality_nc_quantity(open_nc, "qtd_aprovada"))
                                rejected = min(qty - approved, self._quality_nc_quantity(open_nc, "qtd_rejeitada"))
                                if approved > 0 or rejected > 0:
                                    pending_guess = max(0.0, qty - approved - rejected)
                                    status = "EM_INSPECAO" if pending_guess > 0 else ("APROVADO" if approved > 0 else "REJEITADO")
                    pending = max(0.0, qty - approved - rejected)
                    rows.append(
                        {
                            "movement_id": movement_id,
                            "note": note,
                            "line": line,
                            "movement": movement,
                            "note_number": note_number,
                            "line_index": line_index,
                            "movement_index": movement_index,
                            "entity_type": entity_type,
                            "entity_id": entity_id,
                            "qty": qty,
                            "approved_qty": approved,
                            "rejected_qty": rejected,
                            "pending_qty": pending,
                            "status": status,
                        }
                    )
        if changed:
            self._save(force=True, audit=False)
        return rows

    def _quality_sync_pending_from_delivery_movements(self) -> None:
        movement_rows = self._quality_iter_delivery_movements(ensure_ids=True)
        data = self.ensure_data()
        totals: dict[tuple[str, str], dict[str, float]] = {}
        for row in movement_rows:
            key = (str(row.get("entity_type", "") or ""), str(row.get("entity_id", "") or ""))
            bucket = totals.setdefault(key, {"received": 0.0, "pending": 0.0, "approved": 0.0, "rejected": 0.0})
            bucket["received"] += self._parse_float(row.get("qty", 0), 0)
            bucket["pending"] += self._parse_float(row.get("pending_qty", 0), 0)
            bucket["approved"] += self._parse_float(row.get("approved_qty", 0), 0)
            bucket["rejected"] += self._parse_float(row.get("rejected_qty", 0), 0)
        changed = False

        def _apply(item: dict[str, Any], entity_type: str, entity_id: str) -> None:
            nonlocal changed
            total = totals.get((entity_type, entity_id))
            if not total:
                return
            pending = round(total["pending"], 4)
            received = round(total["received"], 4)
            approved = round(total["approved"], 4)
            rejected = round(total["rejected"], 4)
            updates = {
                "quality_pending_qty": pending,
                "quality_received_qty": received,
                "quality_approved_qty": approved,
                "quality_rejected_qty": rejected,
            }
            for key, value in updates.items():
                if abs(self._parse_float(item.get(key, 0), 0) - value) > 1e-6:
                    item[key] = value
                    changed = True
            current_status = self._quality_status_code(item.get("quality_status", item.get("inspection_status", "")))
            if pending > 0:
                target_status = "EM_AVERIGUACAO" if current_status == "EM_AVERIGUACAO" else "EM_INSPECAO"
                target_blocked = True
            elif approved > 0:
                target_status = "APROVADO"
                target_blocked = False
            elif rejected > 0:
                target_status = "REJEITADO"
                target_blocked = True
            else:
                target_status = current_status
                target_blocked = current_status != "APROVADO"
            if str(item.get("quality_status", "") or "").strip() != target_status:
                item["quality_status"] = target_status
                item["inspection_status"] = target_status
                changed = True
            if bool(item.get("quality_blocked")) != target_blocked:
                item["quality_blocked"] = target_blocked
                changed = True

        for material in list(data.get("materiais", []) or []):
            if isinstance(material, dict):
                _apply(material, "Material", str(material.get("id", "") or "").strip())
        for product in list(data.get("produtos", []) or []):
            if isinstance(product, dict):
                _apply(product, "Produto", str(product.get("codigo", "") or "").strip())
        if changed:
            self._save(force=True, audit=False)

    def _quality_reference_key(self, value: Any, fallback: Any = "") -> str:
        raw = str(value or fallback or "").strip()
        match = re.search(r"\bNE-\d{4}-\d{4}\b", raw, flags=re.IGNORECASE)
        if match:
            return match.group(0).upper()
        return re.sub(r"\s+", " ", raw).casefold()

    def _quality_nc_key(self, payload: dict[str, Any]) -> tuple[str, str, str, str]:
        origem_raw = str(payload.get("origem", "") or "").strip()
        origem = re.sub(r"\s+", " ", origem_raw).casefold()
        if "rece" in origem and "fornecedor" in origem:
            origem = "rececao fornecedor"
        referencia = self._quality_reference_key(payload.get("referencia", ""), payload.get("ne_numero", ""))
        entidade_tipo = str(payload.get("entidade_tipo", "") or payload.get("linked_entity_type", "") or "").strip()
        entidade_id = str(payload.get("entidade_id", "") or payload.get("linked_entity_id", "") or "").strip()
        if not entidade_tipo and str(payload.get("material_id", "") or "").strip():
            entidade_tipo = "Material"
            entidade_id = str(payload.get("material_id", "") or "").strip()
        if not entidade_id:
            entidade_id = str(payload.get("material_id", "") or payload.get("produto_codigo", "") or payload.get("fornecedor_id", "") or payload.get("fornecedor_nome", "") or "").strip()
        return (origem, referencia, entidade_tipo.casefold(), entidade_id.casefold())

    def _quality_is_open_nc(self, row: dict[str, Any]) -> bool:
        return str(row.get("estado", "") or "Aberta").strip().casefold() == "aberta"

    def _quality_find_open_nc(self, payload: dict[str, Any], *, exclude_id: str = "") -> dict[str, Any] | None:
        key = self._quality_nc_key(payload)
        exclude = str(exclude_id or "").strip()
        for row in list(self.ensure_data().get("quality_nonconformities", []) or []):
            if not isinstance(row, dict) or not self._quality_is_open_nc(row):
                continue
            if exclude and str(row.get("id", "") or "").strip() == exclude:
                continue
            if self._quality_nc_key(row) == key:
                return row
        return None

    def _quality_nc_quantity(self, row: dict[str, Any] | None, field: str) -> float:
        if not isinstance(row, dict):
            return 0.0
        for key in (field, field.replace("qtd_", "quality_"), field.replace("qtd_", "")):
            if key in row:
                value = self._parse_float(row.get(key, 0), 0)
                if value:
                    return value
        desc = str(row.get("descricao", "") or "")
        label = {
            "qtd_recebida": "recebido",
            "qtd_aprovada": "aprovado",
            "qtd_rejeitada": "rejeitado",
            "qtd_pendente": "pendente",
        }.get(field, field)
        match = re.search(rf"{re.escape(label)}\s*:\s*([0-9]+(?:[.,][0-9]+)?)", desc, flags=re.IGNORECASE)
        if match:
            return self._parse_float(match.group(1), 0)
        return 0.0

    def _quality_normalize_open_nc_duplicates(self) -> None:
        data = self.ensure_data()
        rows = [row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict)]
        first_by_key: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        changed = False
        for row in rows:
            if not self._quality_is_open_nc(row):
                continue
            key = self._quality_nc_key(row)
            if not all(key):
                continue
            keeper = first_by_key.get(key)
            if keeper is None:
                first_by_key[key] = row
                continue
            row["estado"] = "Cancelada"
            row["updated_at"] = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
            row["updated_by"] = self._current_user_label()
            row["acao"] = (
                str(row.get("acao", "") or "").strip()
                + f"\nCancelada automaticamente: NC duplicada de {str(keeper.get('id', '') or '').strip()}."
            ).strip()
            changed = True
        if changed:
            self._save(force=True, audit=False)

    def quality_reception_rows(self, filter_text: str = "", state_filter: str = "Pendentes") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state = str(state_filter or "Pendentes").strip().casefold()
        rows: list[dict[str, Any]] = []
        changed_quarantine = False

        def accept_status(status: str) -> bool:
            code = self._quality_status_code(status)
            if "todo" in state:
                return True
            if "aprov" in state:
                return code == "APROVADO"
            if "rejeit" in state:
                return code == "REJEITADO"
            if "devol" in state:
                return code == "DEVOLVER_FORNECEDOR"
            if "averig" in state:
                return code == "EM_AVERIGUACAO"
            return code == "EM_INSPECAO"

        for movement_row in self._quality_iter_delivery_movements(ensure_ids=True):
            entity_type = str(movement_row.get("entity_type", "") or "").strip()
            entity_id = str(movement_row.get("entity_id", "") or "").strip()
            if not entity_type or not entity_id:
                continue
            status = str(movement_row.get("status", "") or "EM_INSPECAO").strip()
            if not accept_status(status):
                continue
            pending_qty = self._parse_float(movement_row.get("pending_qty", 0), 0)
            if "pend" in state and pending_qty <= 0:
                continue
            target = (
                self.material_by_id(entity_id)
                if entity_type == "Material"
                else next((row for row in list(self.ensure_data().get("produtos", []) or []) if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == entity_id), None)
            )
            target = dict(target or {})
            line = dict(movement_row.get("line", {}) or {})
            movement = dict(movement_row.get("movement", {}) or {})
            row = {
                "tipo": entity_type,
                "id": entity_id,
                "ref": entity_id,
                "movement_id": str(movement_row.get("movement_id", "") or "").strip(),
                "referencia": str(movement_row.get("note_number", "") or "").strip(),
                "material": str(target.get("material", "") or target.get("categoria", "") or line.get("material", "") or line.get("categoria", "") or "").strip(),
                "espessura": str(target.get("espessura", "") or line.get("espessura", "") or "").strip(),
                "descricao": str(target.get("descricao", "") or line.get("descricao", "") or "").strip(),
                "lote": str(movement.get("lote_fornecedor", "") or target.get("lote_fornecedor", "") or "").strip(),
                "fornecedor": str(target.get("inspection_supplier_name", "") or target.get("fornecedor", "") or "").strip(),
                "fornecedor_id": str(target.get("inspection_supplier_id", "") or target.get("fornecedor_id", "") or "").strip(),
                "logistic_status": str(movement.get("logistic_status", "") or target.get("logistic_status", "") or "RECEBIDO").strip(),
                "quality_status": self._quality_status_code(status),
                "defeito": str(movement.get("inspection_defect", "") or target.get("inspection_defect", "") or "").strip(),
                "decisao": str(movement.get("inspection_decision", "") or target.get("inspection_decision", "") or "").strip(),
                "qtd": round(pending_qty, 4),
                "qtd_recebida": round(self._parse_float(movement_row.get("qty", 0), 0), 4),
                "qtd_aprovada": round(self._parse_float(movement_row.get("approved_qty", 0), 0), 4),
                "qtd_rejeitada": round(self._parse_float(movement_row.get("rejected_qty", 0), 0), 4),
                "qtd_disponivel": self._parse_float(target.get("qty" if entity_type == "Produto" else "quantidade", 0), 0),
                "nc_id": str(movement.get("quality_nc_id", "") or target.get("quality_nc_id", "") or target.get("supplier_claim_id", "") or "").strip(),
                "guia": str(movement.get("guia", "") or target.get("inspection_guia", "") or "").strip(),
                "fatura": str(movement.get("fatura", "") or target.get("inspection_fatura", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (str(item.get("quality_status", "")) == "APROVADO", str(item.get("referencia", "")), str(item.get("ref", ""))))
        if changed_quarantine:
            self._save(force=True, audit=False)
        return rows

    def quality_reception_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        self._quality_sync_pending_from_delivery_movements()
        data = self.ensure_data()
        item_type = str(payload.get("tipo", payload.get("kind", "")) or "").strip().casefold()
        item_id = str(payload.get("id", payload.get("ref", "")) or "").strip()
        movement_id = str(payload.get("movement_id", "") or "").strip()
        status = self._quality_status_code(payload.get("quality_status", payload.get("inspection_status", "")))
        defect = str(payload.get("defeito", payload.get("inspection_defect", "")) or "").strip()
        decision = str(payload.get("decisao", payload.get("inspection_decision", "")) or "").strip()
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        if not item_id:
            raise ValueError("Seleciona uma linha de receção para avaliar.")
        if not movement_id:
            raise ValueError("A qualidade só pode avaliar movimentos de receção provenientes de notas de encomenda.")

        movement_ctx = None
        movement_ctx = next(
            (
                row
                for row in self._quality_iter_delivery_movements(ensure_ids=True)
                if str(row.get("movement_id", "") or "").strip() == movement_id
            ),
            None,
        )
        if movement_ctx is None:
            raise ValueError("Movimento de receção não encontrado.")
        item_type = str(movement_ctx.get("entity_type", "") or item_type).casefold()
        item_id = str(movement_ctx.get("entity_id", "") or item_id).strip()

        if "prod" in item_type:
            target = next((row for row in list(data.get("produtos", []) or []) if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == item_id), None)
            entity_type = "Produto"
            entity_id = item_id
            reference = str(payload.get("referencia", "") or (target or {}).get("inspection_note_number", "") or "").strip()
            entity_label = " | ".join(part for part in (entity_id, str((target or {}).get("descricao", "") or "").strip()) if part)
        else:
            target = self.material_by_id(item_id)
            entity_type = "Material"
            entity_id = item_id
            reference = str(payload.get("referencia", "") or (target or {}).get("inspection_note_number", "") or (target or {}).get("origem_ne", "") or "").strip()
            entity_label = self._quality_link_label(entity_type, entity_id)
        if target is None:
            raise ValueError("Linha de receção não encontrada.")

        self._quality_quarantine_pending_stock(
            target,
            kind=entity_type,
            max_qty=self._parse_float((movement_ctx or {}).get("pending_qty", 0), 0) if movement_ctx is not None else None,
        )
        before = copy.deepcopy(target)
        qty_key = "qty" if entity_type == "Produto" else "quantidade"
        pending_qty = (
            self._parse_float((movement_ctx or {}).get("pending_qty", 0), 0)
            if movement_ctx is not None
            else self._parse_float(target.get("quality_pending_qty", 0), 0)
        )
        received_qty = (
            self._parse_float((movement_ctx or {}).get("qty", pending_qty), pending_qty)
            if movement_ctx is not None
            else self._parse_float(target.get("quality_received_qty", pending_qty), pending_qty)
        )
        approved_qty = self._parse_float(payload.get("qtd_aprovada", payload.get("approved_qty", None)), -1)
        rejected_qty = self._parse_float(payload.get("qtd_rejeitada", payload.get("rejected_qty", None)), -1)
        if approved_qty < 0 and rejected_qty < 0:
            approved_qty = pending_qty if status == "APROVADO" else 0.0
            rejected_qty = pending_qty if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} else 0.0
        else:
            approved_qty = max(0.0, approved_qty)
            rejected_qty = max(0.0, rejected_qty)
        if approved_qty + rejected_qty > pending_qty + 1e-9:
            raise ValueError(
                "As quantidades aprovadas/rejeitadas não podem ultrapassar a quantidade pendente "
                f"({self._fmt(pending_qty)})."
            )
        if status == "APROVADO" and rejected_qty <= 0:
            defect = ""
            decision = decision or "Libertar para stock"
        remaining_qty = max(0.0, pending_qty - approved_qty - rejected_qty)
        if status == "APROVADO" and approved_qty <= 0 and pending_qty > 0:
            raise ValueError("Para aprovar, indica a quantidade boa a libertar para stock.")
        if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} and rejected_qty <= 0 and pending_qty > 0:
            raise ValueError("Para rejeitar/devolver, indica a quantidade rejeitada.")
        available_before = self._parse_float(target.get(qty_key, 0), 0)
        target["logistic_status"] = str(target.get("logistic_status", "") or "RECEBIDO").strip()
        effective_status = status
        if remaining_qty > 0:
            effective_status = "EM_AVERIGUACAO" if status == "EM_AVERIGUACAO" else "EM_INSPECAO"
        elif approved_qty > 0:
            effective_status = "APROVADO"
        elif rejected_qty > 0:
            effective_status = status if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} else "REJEITADO"
        target["quality_status"] = effective_status
        target["inspection_status"] = effective_status
        target["inspection_defect"] = defect
        target["inspection_decision"] = decision or ("Libertar para stock" if effective_status == "APROVADO" else "Aguardar decisão da qualidade")
        target["inspection_at"] = now
        target["inspection_by"] = self._current_user_label()
        target["quality_blocked"] = effective_status != "APROVADO"
        target["atualizado_em"] = now
        target["quality_last_received_qty"] = round(received_qty, 4)
        target["quality_last_approved_qty"] = round(approved_qty, 4)
        target["quality_last_rejected_qty"] = round(rejected_qty, 4)

        if movement_ctx is not None:
            movement = movement_ctx["movement"]
            movement["inspection_status"] = effective_status
            movement["quality_status"] = effective_status
            movement["inspection_defect"] = defect
            movement["inspection_decision"] = target["inspection_decision"]
            movement["inspection_at"] = now
            movement["inspection_by"] = self._current_user_label()
            movement["quality_movement_id"] = movement_id
            movement["stock_ref"] = entity_id
            movement["quality_approved_qty"] = self._parse_float(movement_ctx.get("approved_qty", 0), 0) + approved_qty
            movement["quality_rejected_qty"] = self._parse_float(movement_ctx.get("rejected_qty", 0), 0) + rejected_qty
            movement["quality_pending_qty"] = remaining_qty

        nc_payload = {
            "origem": "Receção fornecedor",
            "referencia": reference,
            "entidade_tipo": entity_type,
            "entidade_id": entity_id,
            "entidade_label": entity_label,
            "tipo": "Fornecedor",
            "gravidade": "Alta" if status == "REJEITADO" else "Media",
            "estado": "Aberta",
            "responsavel": "Qualidade",
            "descricao": (
                f"Avaliação de receção marcada como {status}. "
                f"Entidade: {entity_label or entity_id}. "
                f"Recebido: {self._fmt(received_qty)} | aprovado: {self._fmt(approved_qty)} | rejeitado: {self._fmt(rejected_qty)}. "
                f"Defeito/observação: {defect or '-'}."
            ),
            "causa": "A apurar com fornecedor/receção.",
            "acao": target["inspection_decision"],
            "fornecedor_id": str(target.get("inspection_supplier_id", "") or target.get("fornecedor_id", "") or "").strip(),
            "fornecedor_nome": str(target.get("inspection_supplier_name", "") or target.get("fornecedor", "") or "").strip(),
            "material_id": entity_id if entity_type == "Material" else "",
            "produto_codigo": entity_id if entity_type == "Produto" else "",
            "lote_fornecedor": str(target.get("lote_fornecedor", "") or "").strip(),
            "ne_numero": reference,
            "guia": str(target.get("inspection_guia", "") or "").strip(),
            "fatura": str(target.get("inspection_fatura", "") or "").strip(),
            "decisao": target["inspection_decision"],
            "movement_id": movement_id,
            "qtd_recebida": received_qty,
            "qtd_aprovada": approved_qty,
            "qtd_rejeitada": rejected_qty,
            "qtd_pendente": remaining_qty,
        }
        existing_open = self._quality_find_open_nc(nc_payload)
        existing_rejected_total = self._quality_nc_quantity(existing_open, "qtd_rejeitada") if existing_open is not None else 0.0
        if existing_open is not None:
            nc_payload["qtd_recebida"] = max(self._quality_nc_quantity(existing_open, "qtd_recebida"), received_qty)
            nc_payload["qtd_aprovada"] = self._quality_nc_quantity(existing_open, "qtd_aprovada") + approved_qty
            nc_payload["qtd_rejeitada"] = self._quality_nc_quantity(existing_open, "qtd_rejeitada") + rejected_qty
            nc_payload["qtd_pendente"] = remaining_qty
        if approved_qty > 0:
                target[qty_key] = available_before + approved_qty
                target["quality_approved_qty"] = self._parse_float(target.get("quality_approved_qty", 0), 0) + approved_qty
                if entity_type == "Produto":
                    self.desktop_main.add_produto_mov(
                        data,
                        tipo="Entrada",
                        operador=self._current_user_label(),
                        codigo=entity_id,
                        descricao=str(target.get("descricao", "") or "").strip(),
                        qtd=approved_qty,
                        antes=available_before,
                        depois=target[qty_key],
                        obs=f"Aprovado pela qualidade | {reference}",
                        origem="Qualidade",
                        ref_doc=reference,
                    )
                else:
                    self.desktop_main.log_stock(
                        data,
                        "ENTRADA_QUALIDADE",
                        f"{entity_id} qtd={approved_qty} ref={reference}",
                        operador=self._current_user_label(),
                    )
        if effective_status == "APROVADO" and rejected_qty <= 0 and existing_rejected_total <= 0:
            if existing_open is not None:
                self.quality_nc_close(str(existing_open.get("id", "") or ""), target["inspection_decision"])
            if entity_type == "Material":
                target["supplier_claim_id"] = ""
            target["quality_nc_id"] = ""
            self._append_audit_event(data, action="Receção aprovada", entity_type=entity_type, entity_id=entity_id, summary=target["inspection_decision"], before=before, after=target)
        else:
            if rejected_qty > 0 or (existing_rejected_total > 0 and approved_qty > 0) or status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} or defect or bool(payload.get("create_nc")):
                if existing_open is not None:
                    nc_payload["id"] = str(existing_open.get("id", "") or "").strip()
                nc = self.quality_nc_save(nc_payload)
                target["quality_nc_id"] = str(nc.get("id", "") or "").strip()
                if movement_ctx is not None:
                    movement_ctx["movement"]["quality_nc_id"] = target["quality_nc_id"]
                if entity_type == "Material":
                    target["supplier_claim_id"] = target["quality_nc_id"]
            if status == "DEVOLVER_FORNECEDOR" or "devol" in target["inspection_decision"].casefold():
                doc = self._quality_return_document(target, entity_type=entity_type, entity_id=entity_id, reference=reference, nc_id=str(target.get("quality_nc_id", "") or ""))
                if doc:
                    target["quality_return_document_id"] = str(doc.get("id", "") or "").strip()
            self._append_audit_event(data, action="Receção em inspeção", entity_type=entity_type, entity_id=entity_id, summary=target["inspection_decision"], before=before, after=target)

        for note in list(data.get("notas_encomenda", []) or []):
            if not isinstance(note, dict):
                continue
            for line in list(note.get("linhas", []) or []):
                if not isinstance(line, dict):
                    continue
                line_ref = str(line.get("ref", "") or "").strip()
                if line_ref != entity_id:
                    continue
                line["quality_status"] = status
                line["inspection_status"] = status
                line["inspection_defect"] = defect
                line["inspection_decision"] = target["inspection_decision"]
                line["quality_nc_id"] = str(target.get("quality_nc_id", "") or "").strip()
                line["_stock_in"] = effective_status == "APROVADO"
                for movement in list(line.get("entregas_linha", []) or []):
                    if not isinstance(movement, dict):
                        continue
                    movement_matches = (
                        movement_id
                        and str(movement.get("quality_movement_id", "") or "").strip() == movement_id
                    ) or (
                        not movement_id
                        and str(movement.get("stock_ref", "") or line_ref).strip() == entity_id
                    )
                    if movement_matches:
                        movement["quality_status"] = effective_status
                        movement["quality_nc_id"] = line["quality_nc_id"]
                line_movements = [mv for mv in list(line.get("entregas_linha", []) or []) if isinstance(mv, dict)]
                if line_movements:
                    pending_line = sum(self._quality_movement_pending_qty(mv) for mv in line_movements)
                    approved_line = sum(self._parse_float(mv.get("quality_approved_qty", 0), 0) for mv in line_movements)
                    rejected_line = sum(self._parse_float(mv.get("quality_rejected_qty", 0), 0) for mv in line_movements)
                    if pending_line > 0:
                        line_status = "EM_INSPECAO"
                    elif approved_line > 0:
                        line_status = "APROVADO"
                    elif rejected_line > 0:
                        line_status = "REJEITADO"
                    else:
                        line_status = effective_status
                    line["quality_status"] = line_status
                    line["inspection_status"] = line_status
                    line["_stock_in"] = line_status == "APROVADO"
        self._sync_ne_from_materia()
        self._quality_sync_pending_from_delivery_movements()
        self._save(force=True, audit=False)
        return {"tipo": entity_type, "id": entity_id, "quality_status": effective_status, "quality_nc_id": str(target.get("quality_nc_id", "") or "").strip()}

    def _quality_return_document(
        self,
        target: dict[str, Any],
        *,
        entity_type: str,
        entity_id: str,
        reference: str,
        nc_id: str = "",
    ) -> dict[str, Any] | None:
        data = self.ensure_data()
        docs = data.setdefault("quality_documents", [])
        existing_id = str(target.get("quality_return_document_id", "") or "").strip()
        if existing_id:
            existing = next((row for row in docs if isinstance(row, dict) and str(row.get("id", "") or "").strip() == existing_id), None)
            if isinstance(existing, dict):
                return existing
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        doc_id = self._next_prefixed_id(docs, "DEV")
        qty_key = "qty" if entity_type == "Produto" else "quantidade"
        pending_qty = self._parse_float(target.get("quality_pending_qty", 0), 0)
        description = " | ".join(
            part
            for part in (
                f"{entity_type} {entity_id}",
                str(target.get("material", "") or target.get("descricao", "") or "").strip(),
                f"Qtd a devolver {self._fmt(pending_qty)}",
                f"Fornecedor {str(target.get('inspection_supplier_name', '') or target.get('fornecedor', '') or '-').strip()}",
                f"NC {nc_id}" if nc_id else "",
            )
            if part
        )
        doc = {
            "id": doc_id,
            "titulo": f"Nota devolução fornecedor {reference or entity_id}",
            "tipo": "Nota devolução fornecedor",
            "entidade": entity_type,
            "entidade_tipo": entity_type,
            "referencia": reference,
            "entidade_id": entity_id,
            "versao": "1",
            "estado": "Rascunho",
            "responsavel": "Qualidade",
            "caminho": "",
            "obs": description,
            "created_at": now,
            "updated_at": now,
            "created_by": self._current_user_label(),
            "nc_id": nc_id,
            "qtd": pending_qty,
            "qtd_stock": self._parse_float(target.get(qty_key, 0), 0),
        }
        docs.append(doc)
        return doc

    def quality_nc_rows(self, filter_text: str = "", state_filter: str = "Ativas") -> list[dict[str, Any]]:
        self._quality_normalize_open_nc_duplicates()
        query = str(filter_text or "").strip().lower()
        state = str(state_filter or "Ativas").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("quality_nonconformities", []) or []):
            if not isinstance(raw, dict):
                continue
            row = dict(raw)
            estado = str(row.get("estado", "") or "Aberta").strip() or "Aberta"
            estado_norm = estado.lower()
            if state not in {"todos", "todas", "all"}:
                if "ativ" in state and estado_norm in {"fechada", "cancelada"}:
                    continue
                if "abert" in state and estado_norm != "aberta":
                    continue
                if "trat" in state and "trat" not in estado_norm:
                    continue
                if "fech" in state and estado_norm != "fechada":
                    continue
            emitted = {
                "id": str(row.get("id", "") or "").strip(),
                "origem": str(row.get("origem", "") or "").strip(),
                "referencia": str(row.get("referencia", "") or "").strip(),
                "entidade_tipo": str(row.get("entidade_tipo", "") or row.get("linked_entity_type", "") or "").strip(),
                "entidade_id": str(row.get("entidade_id", "") or row.get("linked_entity_id", "") or "").strip(),
                "entidade_label": str(row.get("entidade_label", "") or row.get("linked_entity_label", "") or "").strip(),
                "tipo": str(row.get("tipo", "") or "").strip(),
                "gravidade": str(row.get("gravidade", "") or "Media").strip(),
                "estado": estado,
                "responsavel": str(row.get("responsavel", "") or "").strip(),
                "prazo": str(row.get("prazo", "") or "").strip()[:10],
                "descricao": str(row.get("descricao", "") or "").strip(),
                "causa": str(row.get("causa", "") or "").strip(),
                "acao": str(row.get("acao", "") or "").strip(),
                "eficacia": str(row.get("eficacia", "") or "").strip(),
                "fornecedor_id": str(row.get("fornecedor_id", "") or "").strip(),
                "fornecedor_nome": str(row.get("fornecedor_nome", "") or "").strip(),
                "material_id": str(row.get("material_id", "") or "").strip(),
                "lote_fornecedor": str(row.get("lote_fornecedor", "") or "").strip(),
                "ne_numero": str(row.get("ne_numero", "") or "").strip(),
                "decisao": str(row.get("decisao", "") or "").strip(),
                "movement_id": str(row.get("movement_id", "") or "").strip(),
                "qtd_recebida": round(self._quality_nc_quantity(row, "qtd_recebida"), 4),
                "qtd_aprovada": round(self._quality_nc_quantity(row, "qtd_aprovada"), 4),
                "qtd_rejeitada": round(self._quality_nc_quantity(row, "qtd_rejeitada"), 4),
                "qtd_pendente": round(self._quality_nc_quantity(row, "qtd_pendente"), 4),
                "created_at": str(row.get("created_at", "") or "").strip(),
                "closed_at": str(row.get("closed_at", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in emitted.values()):
                continue
            rows.append(emitted)
        rows.sort(key=lambda item: (str(item.get("estado", "")) == "Fechada", str(item.get("prazo", "") or "9999"), str(item.get("id", ""))), reverse=False)
        return rows

    def quality_nc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        rows = data.setdefault("quality_nonconformities", [])
        nc_id = str(payload.get("id", "") or "").strip()
        existing = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == nc_id), None) if nc_id else None
        before = copy.deepcopy(existing) if isinstance(existing, dict) else None
        if not nc_id:
            nc_id = self._next_prefixed_id(rows, "NC")
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        row = {
            "id": nc_id,
            "origem": str(payload.get("origem", "") or "").strip(),
            "referencia": str(payload.get("referencia", "") or "").strip(),
            "entidade_tipo": str(payload.get("entidade_tipo", payload.get("linked_entity_type", "")) or "").strip(),
            "entidade_id": str(payload.get("entidade_id", payload.get("linked_entity_id", "")) or "").strip(),
            "tipo": str(payload.get("tipo", "") or "Processo").strip() or "Processo",
            "gravidade": str(payload.get("gravidade", "") or "Media").strip() or "Media",
            "estado": str(payload.get("estado", "") or (existing or {}).get("estado", "Aberta") or "Aberta").strip() or "Aberta",
            "responsavel": str(payload.get("responsavel", "") or "").strip(),
            "prazo": str(payload.get("prazo", "") or "").strip()[:10],
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "causa": str(payload.get("causa", "") or "").strip(),
            "acao": str(payload.get("acao", "") or "").strip(),
            "eficacia": str(payload.get("eficacia", "") or "").strip(),
            "fornecedor_id": str(payload.get("fornecedor_id", (existing or {}).get("fornecedor_id", "")) or "").strip(),
            "fornecedor_nome": str(payload.get("fornecedor_nome", (existing or {}).get("fornecedor_nome", "")) or "").strip(),
            "material_id": str(payload.get("material_id", (existing or {}).get("material_id", "")) or "").strip(),
            "lote_fornecedor": str(payload.get("lote_fornecedor", (existing or {}).get("lote_fornecedor", "")) or "").strip(),
            "ne_numero": str(payload.get("ne_numero", (existing or {}).get("ne_numero", "")) or "").strip(),
            "guia": str(payload.get("guia", (existing or {}).get("guia", "")) or "").strip(),
            "fatura": str(payload.get("fatura", (existing or {}).get("fatura", "")) or "").strip(),
            "decisao": str(payload.get("decisao", (existing or {}).get("decisao", "")) or "").strip(),
            "movement_id": str(payload.get("movement_id", (existing or {}).get("movement_id", "")) or "").strip(),
            "qtd_recebida": round(self._parse_float(payload.get("qtd_recebida", (existing or {}).get("qtd_recebida", 0)), 0), 4),
            "qtd_aprovada": round(self._parse_float(payload.get("qtd_aprovada", (existing or {}).get("qtd_aprovada", 0)), 0), 4),
            "qtd_rejeitada": round(self._parse_float(payload.get("qtd_rejeitada", (existing or {}).get("qtd_rejeitada", 0)), 0), 4),
            "qtd_pendente": round(self._parse_float(payload.get("qtd_pendente", (existing or {}).get("qtd_pendente", 0)), 0), 4),
            "created_at": str((existing or {}).get("created_at", "") or now),
            "updated_at": now,
            "created_by": str((existing or {}).get("created_by", "") or self._current_user_label()),
            "updated_by": self._current_user_label(),
            "closed_at": str((existing or {}).get("closed_at", "") or "").strip(),
        }
        row["entidade_label"] = str(payload.get("entidade_label", "") or "").strip() or self._quality_link_label(
            row["entidade_tipo"], row["entidade_id"]
        )
        if not row["referencia"] and row["entidade_id"]:
            row["referencia"] = row["entidade_id"]
        if self._quality_is_open_nc(row):
            duplicate = self._quality_find_open_nc(row, exclude_id=nc_id)
            if duplicate is not None:
                dup_id = str(duplicate.get("id", "") or "").strip()
                raise ValueError(
                    f"Já existe uma NC aberta ({dup_id}) para esta origem, referência e entidade. "
                    "Fecha ou edita essa NC antes de criar outra."
                )
        if existing is None:
            rows.append(row)
        else:
            existing.update(row)
            row = existing
        self._append_audit_event(
            data,
            action="NC guardada",
            entity_type="Nao conformidade",
            entity_id=nc_id,
            summary=f"{row.get('tipo', '')} | {row.get('estado', '')} | {row.get('referencia', '')}",
            before=before,
            after=row,
        )
        self._save(force=True, audit=False)
        return dict(row)

    def quality_nc_close(self, nc_id: str, eficacia: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        target = next((row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict) and str(row.get("id", "") or "").strip() == str(nc_id or "").strip()), None)
        if target is None:
            raise ValueError("Nao conformidade nao encontrada.")
        before = copy.deepcopy(target)
        target["estado"] = "Fechada"
        target["closed_at"] = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        target["closed_by"] = self._current_user_label()
        if str(eficacia or "").strip():
            target["eficacia"] = str(eficacia or "").strip()
        self._append_audit_event(data, action="NC fechada", entity_type="Nao conformidade", entity_id=str(nc_id), summary=str(target.get("eficacia", "") or ""), before=before, after=target)
        self._save(force=True, audit=False)
        return dict(target)

    def quality_nc_release_material(self, nc_id: str, decision: str = "Aprovado pela qualidade") -> dict[str, Any]:
        data = self.ensure_data()
        nc_id_txt = str(nc_id or "").strip()
        target = next((row for row in list(data.get("quality_nonconformities", []) or []) if isinstance(row, dict) and str(row.get("id", "") or "").strip() == nc_id_txt), None)
        if target is None:
            raise ValueError("Nao conformidade nao encontrada.")
        material_id = str(target.get("material_id", "") or "").strip()
        if not material_id and str(target.get("entidade_tipo", "") or "").strip() == "Material":
            material_id = str(target.get("entidade_id", "") or "").strip()
        if not material_id:
            raise ValueError("Esta NC nao esta ligada a um material.")
        material = self.material_by_id(material_id)
        if material is None:
            raise ValueError("Material ligado a NC nao encontrado.")
        before_material = copy.deepcopy(material)
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        self._quality_quarantine_pending_stock(material, kind="Material")
        pending_qty = self._parse_float(material.get("quality_pending_qty", 0), 0)
        before_qty = self._parse_float(material.get("quantidade", 0), 0)
        if pending_qty > 0:
            material["quantidade"] = before_qty + pending_qty
            material["quality_pending_qty"] = 0.0
            material["quality_approved_qty"] = self._parse_float(material.get("quality_approved_qty", 0), 0) + pending_qty
            self.desktop_main.log_stock(
                data,
                "ENTRADA_QUALIDADE",
                f"{material_id} qtd={pending_qty} NC={nc_id_txt}",
                operador=self._current_user_label(),
            )
        material["quality_status"] = "APROVADO"
        material["inspection_status"] = "APROVADO"
        material["quality_blocked"] = False
        material["inspection_decision"] = str(decision or "Aprovado pela qualidade").strip()
        material["quality_nc_id"] = ""
        material["supplier_claim_id"] = ""
        material["quality_released_at"] = now
        material["quality_released_by"] = self._current_user_label()
        material["atualizado_em"] = now
        target["decisao"] = str(decision or "Aprovado pela qualidade").strip()
        target["acao"] = (str(target.get("acao", "") or "").strip() + f"\nLibertacao de material: {material['inspection_decision']}").strip()
        target["estado"] = "Fechada"
        target["closed_at"] = now
        target["closed_by"] = self._current_user_label()
        target["updated_at"] = now
        target["updated_by"] = self._current_user_label()
        self._append_audit_event(
            data,
            action="Material libertado pela qualidade",
            entity_type="Material",
            entity_id=material_id,
            summary=f"NC {nc_id_txt}: {material['inspection_decision']}",
            before=before_material,
            after=material,
        )
        self._sync_ne_from_materia()
        self._save(force=True, audit=False)
        return {"material_id": material_id, "quality_status": "APROVADO", "nc_id": nc_id_txt}

    def quality_nc_remove(self, nc_id: str) -> None:
        data = self.ensure_data()
        value = str(nc_id or "").strip()
        rows = list(data.get("quality_nonconformities", []) or [])
        before = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == value), None)
        data["quality_nonconformities"] = [row for row in rows if not (isinstance(row, dict) and str(row.get("id", "") or "").strip() == value)]
        if before is None:
            raise ValueError("Nao conformidade nao encontrada.")
        self._append_audit_event(data, action="NC removida", entity_type="Nao conformidade", entity_id=value, summary=str(before.get("descricao", "") or ""), before=before)
        self._save(force=True, audit=False)

    def quality_document_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("quality_documents", []) or []):
            if not isinstance(raw, dict):
                continue
            row = {
                "id": str(raw.get("id", "") or "").strip(),
                "titulo": str(raw.get("titulo", "") or "").strip(),
                "tipo": str(raw.get("tipo", "") or "").strip(),
                "entidade": str(raw.get("entidade", "") or "").strip(),
                "referencia": str(raw.get("referencia", "") or "").strip(),
                "versao": str(raw.get("versao", "") or "").strip(),
                "estado": str(raw.get("estado", "") or "Ativo").strip(),
                "responsavel": str(raw.get("responsavel", "") or "").strip(),
                "caminho": str(raw.get("caminho", "") or "").strip(),
                "obs": str(raw.get("obs", "") or "").strip(),
                "updated_at": str(raw.get("updated_at", "") or raw.get("created_at", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (str(item.get("tipo", "")), str(item.get("titulo", ""))))
        return rows

    def quality_document_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        rows = data.setdefault("quality_documents", [])
        doc_id = str(payload.get("id", "") or "").strip()
        existing = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == doc_id), None) if doc_id else None
        before = copy.deepcopy(existing) if isinstance(existing, dict) else None
        if not doc_id:
            doc_id = self._next_prefixed_id(rows, "DOC")
        titulo = str(payload.get("titulo", "") or "").strip()
        if not titulo:
            raise ValueError("Titulo do documento obrigatorio.")
        source_path = str(payload.get("caminho", "") or "").strip()
        stored_path = source_path
        if source_path:
            stored_path = self._store_shared_file(source_path, "quality/documents", preferred_name=self._file_reference_name(source_path, titulo or doc_id))
        now = str(self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))
        row = {
            "id": doc_id,
            "titulo": titulo,
            "tipo": str(payload.get("tipo", "") or "Evidencia").strip() or "Evidencia",
            "entidade": str(payload.get("entidade", "") or "").strip(),
            "referencia": str(payload.get("referencia", "") or "").strip(),
            "entidade_tipo": str(payload.get("entidade_tipo", payload.get("entidade", "")) or "").strip(),
            "entidade_id": str(payload.get("entidade_id", payload.get("referencia", "")) or "").strip(),
            "versao": str(payload.get("versao", "") or "1").strip() or "1",
            "estado": str(payload.get("estado", "") or "Ativo").strip() or "Ativo",
            "responsavel": str(payload.get("responsavel", "") or "").strip(),
            "caminho": stored_path,
            "obs": str(payload.get("obs", "") or "").strip(),
            "created_at": str((existing or {}).get("created_at", "") or now),
            "updated_at": now,
            "created_by": str((existing or {}).get("created_by", "") or self._current_user_label()),
            "updated_by": self._current_user_label(),
        }
        if existing is None:
            rows.append(row)
        else:
            existing.update(row)
            row = existing
        self._append_audit_event(data, action="Documento qualidade guardado", entity_type="Documento", entity_id=doc_id, summary=titulo, before=before, after=row)
        self._save(force=True, audit=False)
        return dict(row)

    def quality_document_remove(self, doc_id: str) -> None:
        data = self.ensure_data()
        value = str(doc_id or "").strip()
        rows = list(data.get("quality_documents", []) or [])
        before = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == value), None)
        data["quality_documents"] = [row for row in rows if not (isinstance(row, dict) and str(row.get("id", "") or "").strip() == value)]
        if before is None:
            raise ValueError("Documento nao encontrado.")
        self._append_audit_event(data, action="Documento qualidade removido", entity_type="Documento", entity_id=value, summary=str(before.get("titulo", "") or ""), before=before)
        self._save(force=True, audit=False)
