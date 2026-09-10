"""Prepare a reception assessment, NC, return document and stock events."""
import copy
import math
from dataclasses import dataclass
from typing import Any, Callable, Protocol
from .stock_policy import status_code
from .nonconformities import Nonconformities, NonconformityRules
from .delivery_movements import DeliveryMovements

@dataclass
class ReceptionCatalogs:
    materials: list
    products: list
    notes: list
    nonconformities: list
    documents: list

class ReceptionRepository(Protocol):
    def load(self) -> ReceptionCatalogs: ...
    def replace(self, catalogs: ReceptionCatalogs, *, expected: ReceptionCatalogs,
                audit_events: list, stock_events: list, product_events: list) -> None: ...

@dataclass(frozen=True)
class ReceptionRules:
    parse_float: Callable
    now_iso: Callable
    actor: Callable
    next_id: Callable
    format_number: Callable
    is_material: Callable
    quarantine: Callable
    sync_notes: Callable

class PreparedNonconformities:
    """Local NC port; collects events without publishing or persisting."""
    def __init__(self, catalogs, audit_events):
        self.catalogs = catalogs
        self.audit_events = audit_events

    def rows(self):
        return copy.deepcopy(self.catalogs.nonconformities)

    def replace(self, rows, *, expected, event=None):
        if self.catalogs.nonconformities != expected:
            raise ValueError("NC alterada durante a preparacao.")
        self.catalogs.nonconformities = copy.deepcopy(rows)
        if event:
            self.audit_events.append(copy.deepcopy(event))

class ReceptionAssessment:
    def __init__(self, data, rules):
        self.rules = rules
        self.audit_events, self.stock_events, self.product_events = [], [], []
        self.nonconformities = Nonconformities(PreparedNonconformities(data, self.audit_events),
            NonconformityRules(rules.parse_float, rules.now_iso, rules.actor, rules.next_id,
                              lambda kind, identifier: identifier))
        self.movements = DeliveryMovements(rules.parse_float, rules.is_material, self.nonconformities)

    def audit(self, **event):
        self.audit_events.append(copy.deepcopy(event))

    def prepare(self, data: ReceptionCatalogs, payload: dict[str, Any]) -> dict[str, Any]:
        self.movements.reconcile(data.notes, data.materials, data.products)
        item_type = str(payload.get("tipo", payload.get("kind", "")) or "").strip().casefold()
        item_id = str(payload.get("id", payload.get("ref", "")) or "").strip()
        movement_id = str(payload.get("movement_id", "") or "").strip()
        status = status_code(payload.get("quality_status", payload.get("inspection_status", "")))
        defect = str(payload.get("defeito", payload.get("inspection_defect", "")) or "").strip()
        decision = str(payload.get("decisao", payload.get("inspection_decision", "")) or "").strip()
        now = self.rules.now_iso()
        if not item_id:
            raise ValueError("Seleciona uma linha de receção para avaliar.")
        if not movement_id:
            raise ValueError("A qualidade só pode avaliar movimentos de receção provenientes de notas de encomenda.")

        movement_ctx = None
        movement_ctx = next(
            (
                row
                for row in self.movements.rows(data.notes)
                if str(row.get("movement_id", "") or "").strip() == movement_id
            ),
            None,
        )
        if movement_ctx is None:
            raise ValueError("Movimento de receção não encontrado.")
        item_type = str(movement_ctx.get("entity_type", "") or item_type).casefold()
        item_id = str(movement_ctx.get("entity_id", "") or item_id).strip()

        if "prod" in item_type:
            target = next((row for row in data.products if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == item_id), None)
            entity_type = "Produto"
            entity_id = item_id
            reference = str(payload.get("referencia", "") or (target or {}).get("inspection_note_number", "") or "").strip()
            entity_label = " | ".join(part for part in (entity_id, str((target or {}).get("descricao", "") or "").strip()) if part)
        else:
            target = next((row for row in data.materials if str(row.get("id", "")).strip() == item_id), None)
            entity_type = "Material"
            entity_id = item_id
            reference = str(payload.get("referencia", "") or (target or {}).get("inspection_note_number", "") or (target or {}).get("origem_ne", "") or "").strip()
            entity_label = " | ".join(str((target or {}).get(key, "") or "").strip() for key in ("id", "material", "espessura", "formato")).strip(" |")
        if target is None:
            raise ValueError("Linha de receção não encontrada.")

        reference = str(movement_ctx["note_number"])
        self.rules.quarantine(
            target,
            kind=entity_type,
            max_qty=self.rules.parse_float((movement_ctx or {}).get("pending_qty", 0), 0) if movement_ctx is not None else None,
        )
        before = copy.deepcopy(target)
        qty_key = "qty" if entity_type == "Produto" else "quantidade"
        pending_qty = (
            self.rules.parse_float((movement_ctx or {}).get("pending_qty", 0), 0)
            if movement_ctx is not None
            else self.rules.parse_float(target.get("quality_pending_qty", 0), 0)
        )
        received_qty = (
            self.rules.parse_float((movement_ctx or {}).get("qty", pending_qty), pending_qty)
            if movement_ctx is not None
            else self.rules.parse_float(target.get("quality_received_qty", pending_qty), pending_qty)
        )
        approved_qty = self.rules.parse_float(payload.get("qtd_aprovada", payload.get("approved_qty", None)), -1)
        rejected_qty = self.rules.parse_float(payload.get("qtd_rejeitada", payload.get("rejected_qty", None)), -1)
        if not all(math.isfinite(value) for value in (pending_qty, received_qty, approved_qty, rejected_qty)):
            raise ValueError("As quantidades devem ser numeros finitos.")
        if approved_qty < 0 and rejected_qty < 0:
            approved_qty = pending_qty if status == "APROVADO" else 0.0
            rejected_qty = pending_qty if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} else 0.0
        else:
            approved_qty = max(0.0, approved_qty)
            rejected_qty = max(0.0, rejected_qty)
        if approved_qty + rejected_qty > pending_qty + 1e-9:
            raise ValueError(
                "As quantidades aprovadas/rejeitadas não podem ultrapassar a quantidade pendente "
                f"({self.rules.format_number(pending_qty)})."
            )
        if status == "APROVADO" and rejected_qty <= 0:
            defect = ""
            decision = decision or "Libertar para stock"
        remaining_qty = max(0.0, pending_qty - approved_qty - rejected_qty)
        if status == "APROVADO" and approved_qty <= 0 and pending_qty > 0:
            raise ValueError("Para aprovar, indica a quantidade boa a libertar para stock.")
        if status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} and rejected_qty <= 0 and pending_qty > 0:
            raise ValueError("Para rejeitar/devolver, indica a quantidade rejeitada.")
        available_before = self.rules.parse_float(target.get(qty_key, 0), 0)
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
        target["inspection_by"] = self.rules.actor()
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
            movement["inspection_by"] = self.rules.actor()
            movement["quality_movement_id"] = movement_id
            movement["stock_ref"] = entity_id
            movement["quality_approved_qty"] = self.rules.parse_float(movement_ctx.get("approved_qty", 0), 0) + approved_qty
            movement["quality_rejected_qty"] = self.rules.parse_float(movement_ctx.get("rejected_qty", 0), 0) + rejected_qty
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
                f"Recebido: {self.rules.format_number(received_qty)} | aprovado: {self.rules.format_number(approved_qty)} | rejeitado: {self.rules.format_number(rejected_qty)}. "
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
        existing_open = self.nonconformities.find_open(nc_payload)
        existing_rejected_total = self.nonconformities.quantity(existing_open, "qtd_rejeitada") if existing_open is not None else 0.0
        if existing_open is not None:
            nc_payload["qtd_recebida"] = max(self.nonconformities.quantity(existing_open, "qtd_recebida"), received_qty)
            nc_payload["qtd_aprovada"] = self.nonconformities.quantity(existing_open, "qtd_aprovada") + approved_qty
            nc_payload["qtd_rejeitada"] = self.nonconformities.quantity(existing_open, "qtd_rejeitada") + rejected_qty
            nc_payload["qtd_pendente"] = remaining_qty
        if approved_qty > 0:
                target[qty_key] = available_before + approved_qty
                target["quality_approved_qty"] = self.rules.parse_float(target.get("quality_approved_qty", 0), 0) + approved_qty
                if entity_type == "Produto":
                    self.product_events.append(dict(
                        tipo="Entrada",
                        operador=self.rules.actor(),
                        codigo=entity_id,
                        descricao=str(target.get("descricao", "") or "").strip(),
                        qtd=approved_qty,
                        antes=available_before,
                        depois=target[qty_key],
                        obs=f"Aprovado pela qualidade | {reference}",
                        origem="Qualidade",
                        ref_doc=reference,
                    ))
                else:
                    self.stock_events.append(("ENTRADA_QUALIDADE", f"{entity_id} qtd={approved_qty} ref={reference}", self.rules.actor()))
        if effective_status == "APROVADO" and rejected_qty <= 0 and existing_rejected_total <= 0:
            if existing_open is not None:
                self.nonconformities.close(str(existing_open.get("id", "") or ""), target["inspection_decision"])
            if entity_type == "Material":
                target["supplier_claim_id"] = ""
            target["quality_nc_id"] = ""
            self.audit(action="Receção aprovada", entity_type=entity_type, entity_id=entity_id, summary=target["inspection_decision"], before=before, after=target)
        else:
            if rejected_qty > 0 or (existing_rejected_total > 0 and approved_qty > 0) or status in {"REJEITADO", "DEVOLVER_FORNECEDOR"} or defect or bool(payload.get("create_nc")):
                if existing_open is not None:
                    nc_payload["id"] = str(existing_open.get("id", "") or "").strip()
                nc = self.nonconformities.save(nc_payload)
                target["quality_nc_id"] = str(nc.get("id", "") or "").strip()
                if movement_ctx is not None:
                    movement_ctx["movement"]["quality_nc_id"] = target["quality_nc_id"]
                if entity_type == "Material":
                    target["supplier_claim_id"] = target["quality_nc_id"]
            if status == "DEVOLVER_FORNECEDOR" or "devol" in target["inspection_decision"].casefold():
                doc = self.return_document(data, target, entity_type=entity_type, entity_id=entity_id, reference=reference, nc_id=str(target.get("quality_nc_id", "") or ""), return_qty=rejected_qty)
                if doc:
                    target["quality_return_document_id"] = str(doc.get("id", "") or "").strip()
            self.audit(action="Receção em inspeção", entity_type=entity_type, entity_id=entity_id, summary=target["inspection_decision"], before=before, after=target)

        for note in data.notes:
            if not isinstance(note, dict):
                continue
            for line in list(note.get("linhas", []) or []):
                if not isinstance(line, dict):
                    continue
                line_ref = str(line.get("ref", "") or "").strip()
                if line is not movement_ctx["line"]:
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
                    pending_line = sum(self.movements.pending_quantity(mv) for mv in line_movements)
                    approved_line = sum(self.rules.parse_float(mv.get("quality_approved_qty", 0), 0) for mv in line_movements)
                    rejected_line = sum(self.rules.parse_float(mv.get("quality_rejected_qty", 0), 0) for mv in line_movements)
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
        self.rules.sync_notes(data.notes, data.materials)
        self.movements.reconcile(data.notes, data.materials, data.products)
        return {"tipo": entity_type, "id": entity_id, "quality_status": effective_status, "quality_nc_id": str(target.get("quality_nc_id", "") or "").strip()}

    def return_document(
        self,
        data: ReceptionCatalogs,
        target: dict[str, Any],
        *,
        entity_type: str,
        entity_id: str,
        reference: str,
        nc_id: str = "",
        return_qty: float | None = None,
    ) -> dict[str, Any] | None:
        docs = data.documents
        existing_id = str(target.get("quality_return_document_id", "") or "").strip()
        if existing_id:
            existing = next((row for row in docs if isinstance(row, dict) and str(row.get("id", "") or "").strip() == existing_id), None)
            if isinstance(existing, dict):
                return existing
        now = self.rules.now_iso()
        doc_id = self.rules.next_id(docs, "DEV")
        qty_key = "qty" if entity_type == "Produto" else "quantidade"
        pending_qty = self.rules.parse_float(target.get("quality_pending_qty", 0), 0) if return_qty is None else return_qty
        description = " | ".join(
            part
            for part in (
                f"{entity_type} {entity_id}",
                str(target.get("material", "") or target.get("descricao", "") or "").strip(),
                f"Qtd a devolver {self.rules.format_number(pending_qty)}",
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
            "created_by": self.rules.actor(),
            "nc_id": nc_id,
            "qtd": pending_qty,
            "qtd_stock": self.rules.parse_float(target.get(qty_key, 0), 0),
        }
        docs.append(doc)
        return doc


class Receptions:
    def __init__(self, repository: ReceptionRepository, rules: ReceptionRules):
        self.repository, self.rules = repository, rules

    def save(self, payload):
        expected = self.repository.load()
        catalogs = copy.deepcopy(expected)
        assessment = ReceptionAssessment(catalogs, self.rules)
        result = assessment.prepare(catalogs, copy.deepcopy(payload))
        self.repository.replace(catalogs, expected=expected, audit_events=assessment.audit_events,
                                stock_events=assessment.stock_events, product_events=assessment.product_events)
        return result

    def reconcile(self):
        expected = self.repository.load()
        catalogs = copy.deepcopy(expected)
        assessment = ReceptionAssessment(catalogs, self.rules)
        if assessment.movements.reconcile(catalogs.notes, catalogs.materials, catalogs.products):
            self.repository.replace(catalogs, expected=expected, audit_events=[], stock_events=[], product_events=[])

    def rows(self, filter_text: str = "", state_filter: str = "Pendentes") -> list[dict[str, Any]]:
        data = self.repository.load()
        assessment = ReceptionAssessment(data, self.rules)
        query = str(filter_text or "").strip().lower()
        state = str(state_filter or "Pendentes").strip().casefold()
        rows: list[dict[str, Any]] = []

        def accept_status(status: str) -> bool:
            code = status_code(status)
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

        for movement_row in assessment.movements.rows(data.notes):
            entity_type = str(movement_row.get("entity_type", "") or "").strip()
            entity_id = str(movement_row.get("entity_id", "") or "").strip()
            if not entity_type or not entity_id:
                continue
            status = str(movement_row.get("status", "") or "EM_INSPECAO").strip()
            if not accept_status(status):
                continue
            pending_qty = self.rules.parse_float(movement_row.get("pending_qty", 0), 0)
            if "pend" in state and pending_qty <= 0:
                continue
            target = (
                next((row for row in data.materials if str(row.get("id", "")).strip() == entity_id), None)
                if entity_type == "Material"
                else next((row for row in data.products if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == entity_id), None)
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
                "quality_status": status_code(status),
                "defeito": str(movement.get("inspection_defect", "") or target.get("inspection_defect", "") or "").strip(),
                "decisao": str(movement.get("inspection_decision", "") or target.get("inspection_decision", "") or "").strip(),
                "qtd": round(pending_qty, 4),
                "qtd_recebida": round(self.rules.parse_float(movement_row.get("qty", 0), 0), 4),
                "qtd_aprovada": round(self.rules.parse_float(movement_row.get("approved_qty", 0), 0), 4),
                "qtd_rejeitada": round(self.rules.parse_float(movement_row.get("rejected_qty", 0), 0), 4),
                "qtd_disponivel": self.rules.parse_float(target.get("qty" if entity_type == "Produto" else "quantidade", 0), 0),
                "nc_id": str(movement.get("quality_nc_id", "") or target.get("quality_nc_id", "") or target.get("supplier_claim_id", "") or "").strip(),
                "guia": str(movement.get("guia", "") or target.get("inspection_guia", "") or "").strip(),
                "fatura": str(movement.get("fatura", "") or target.get("inspection_fatura", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (str(item.get("quality_status", "")) == "APROVADO", str(item.get("referencia", "")), str(item.get("ref", ""))))
        return rows


    def movement_rows(self):
        data = self.repository.load()
        return ReceptionAssessment(data, self.rules).movements.rows(data.notes)

    def return_document(self, target, **kwargs):
        expected = self.repository.load()
        data = copy.deepcopy(expected)
        assessment = ReceptionAssessment(data, self.rules)
        result = assessment.return_document(data, copy.deepcopy(target), **kwargs)
        if data != expected:
            self.repository.replace(data, expected=expected, audit_events=[], stock_events=[], product_events=[])
        return copy.deepcopy(result)
