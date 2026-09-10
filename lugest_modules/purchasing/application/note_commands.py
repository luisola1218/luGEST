"""Prepare a purchase note and affected price catalogs before one commit."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol
from .pricing import PurchasePricing

@dataclass
class PurchaseCatalogs:
    notes: list[dict[str, Any]]
    materials: list[dict[str, Any]]
    products: list[dict[str, Any]]
    assemblies: list[dict[str, Any]]

class PurchaseWriteRepository(Protocol):
    def load(self) -> PurchaseCatalogs: ...
    def allocate_number(self) -> str: ...
    def save_catalogs(self, catalogs: PurchaseCatalogs, *, expected: PurchaseCatalogs) -> None: ...

@dataclass(frozen=True)
class PurchaseWriteRules:
    normalize_line: Callable
    normalize_supplier: Callable
    supplier: Callable
    now_iso: Callable
    parse_float: Callable
    kind: Callable
    recalculate: Callable

class NoteCommands:
    def __init__(self, repository: PurchaseWriteRepository, rules: PurchaseWriteRules, pricing: PurchasePricing):
        self.repository = repository
        self.rules = rules
        self.pricing = pricing

    def save(self, payload: dict[str, Any]) -> dict[str, Any]:
        payload = deepcopy(payload)
        original = self.repository.load()
        catalogs = deepcopy(original)
        numero = str(payload.get("numero", "") or "").strip()
        fornecedor_id = str(payload.get("fornecedor_id", "") or "").strip()
        fornecedor = str(payload.get("fornecedor", "") or "").strip()
        contacto = str(payload.get("contacto", "") or "").strip()
        referencias_orcamento = str(payload.get("referencias_orcamento", "") or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        obs = str(payload.get("obs", "") or "").strip()
        local_descarga = str(payload.get("local_descarga", "") or "").strip()
        meio_transporte = str(payload.get("meio_transporte", "") or "").strip()
        lines_payload = list(payload.get("lines", []) or [])
        normalized_lines = [self.rules.normalize_line(line) for line in lines_payload]
        resolved_supplier_id, resolved_supplier_text, resolved_contact = self.rules.normalize_supplier(fornecedor_id, fornecedor)
        fornecedor_id = resolved_supplier_id
        fornecedor = resolved_supplier_text
        if not contacto and resolved_contact:
            contacto = resolved_contact
        for line in normalized_lines:
            line_supplier = str(line.get("fornecedor_linha", "") or "").strip()
            if not line_supplier:
                continue
            _, resolved_line_supplier, _ = self.rules.normalize_supplier("", line_supplier)
            if resolved_line_supplier:
                line["fornecedor_linha"] = resolved_line_supplier
        lst = catalogs.notes
        existing = next((row for row in lst if str(row.get("numero", "") or "").strip() == numero), None)
        old_lines = list(existing.get("linhas", []) or []) if isinstance(existing, dict) else []
        if not fornecedor and normalized_lines:
            unique_suppliers = {
                str(line.get("fornecedor_linha", "") or "").strip()
                for line in normalized_lines
                if str(line.get("fornecedor_linha", "") or "").strip()
            }
            if len(unique_suppliers) == 1:
                fornecedor = next(iter(unique_suppliers))
                fornecedor_id, fornecedor, inferred_contact = self.rules.supplier(fornecedor)
                if not contacto:
                    contacto = inferred_contact
        note = {
            "numero": numero,
            "fornecedor": fornecedor,
            "fornecedor_id": fornecedor_id,
            "contacto": contacto,
            "created_at": str(
                (existing or {}).get("created_at", "")
                or payload.get("created_at", "")
                or self.rules.now_iso()
            ).strip(),
            "referencias_orcamento": referencias_orcamento,
            "data_entrega": data_entrega,
            "obs": obs,
            "local_descarga": local_descarga,
            "meio_transporte": meio_transporte,
            "linhas": normalized_lines,
            "estado": str((existing or {}).get("estado", "Em edicao") or "Em edicao").strip() or "Em edicao",
            "oculta": bool((existing or {}).get("oculta", False)),
            "_draft": False,
            "origem_cotacao": str((existing or {}).get("origem_cotacao", "") or "").strip(),
            "ne_geradas": list((existing or {}).get("ne_geradas", []) or []),
            "entregas": list((existing or {}).get("entregas", []) or []),
            "documentos": list((existing or {}).get("documentos", []) or []),
            "guia_ultima": str((existing or {}).get("guia_ultima", "") or "").strip(),
            "fatura_ultima": str((existing or {}).get("fatura_ultima", "") or "").strip(),
            "fatura_caminho_ultima": str((existing or {}).get("fatura_caminho_ultima", "") or "").strip(),
            "data_doc_ultima": str((existing or {}).get("data_doc_ultima", "") or "").strip(),
            "data_ultima_entrega": str((existing or {}).get("data_ultima_entrega", "") or "").strip(),
            "data_entregue": str((existing or {}).get("data_entregue", "") or "").strip(),
            "data_aprovacao": str((existing or {}).get("data_aprovacao", "") or "").strip(),
        }
        for index, line in enumerate(note["linhas"]):
            if index >= len(old_lines):
                if line.get("entregue"):
                    line["qtd_entregue"] = self.rules.parse_float(line.get("qtd", 0), 0)
                continue
            old = old_lines[index]
            qtd_tot = self.rules.parse_float(line.get("qtd", 0), 0)
            qtd_old = self.rules.parse_float(old.get("qtd_entregue", old.get("qtd", 0) if old.get("entregue") else 0), 0)
            qtd_old = max(0.0, min(qtd_tot, qtd_old))
            line["qtd_entregue"] = qtd_old
            if old.get("entregue") or (qtd_tot > 0 and qtd_old >= (qtd_tot - 1e-9)):
                line["entregue"] = True
            if old.get("_stock_in") and qtd_old > 0:
                line["_stock_in"] = True
            for key in ("guia_entrega", "fatura_entrega", "data_doc_entrega", "data_entrega_real", "obs_entrega", "entregas_linha"):
                if old.get(key):
                    line[key] = old.get(key)
        note_kind = self.rules.kind(note)
        if note_kind == "rfq":
            note["fornecedor"] = ""
            note["fornecedor_id"] = ""
            note["contacto"] = ""
        elif not str(note.get("fornecedor_id", "") or "").strip() and str(note.get("fornecedor", "") or "").strip():
            raise ValueError("Seleciona um fornecedor válido da ficha de fornecedores.")
        self.pricing.apply(note["linhas"], catalogs)
        if not numero:
            numero = str(self.repository.allocate_number())
            note["numero"] = numero
        self.rules.recalculate(note)
        if existing:
            existing.update(note)
        else:
            lst.append(note)
        self.repository.save_catalogs(catalogs, expected=original)
        return deepcopy(note)
