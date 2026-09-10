"""Purchase note lifecycle and supplier conversion on prepared records."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class NoteRepository(Protocol):
    def rows(self) -> list[dict[str, Any]]: ...
    def allocate_number(self) -> str: ...
    def replace(self, rows: list[dict[str, Any]], *, expected: list[dict[str, Any]]) -> None: ...

@dataclass(frozen=True)
class NoteRules:
    now_iso: Callable
    kind: Callable
    supplier: Callable
    recalculate: Callable

class NoteLifecycle:
    def __init__(self, repository: NoteRepository, rules: NoteRules):
        self.repository = repository
        self.rules = rules

    def create_draft(self) -> dict[str, Any]:
        original = self.repository.rows()
        notes = deepcopy(original)
        numero = str(self.repository.allocate_number())
        note = {
            "numero": numero,
            "fornecedor": "",
            "fornecedor_id": "",
            "contacto": "",
            "created_at": self.rules.now_iso(),
            "referencias_orcamento": "",
            "data_entrega": "",
            "obs": "",
            "local_descarga": "",
            "meio_transporte": "",
            "linhas": [],
            "total": 0.0,
            "estado": "Em edicao",
            "oculta": False,
            "_draft": True,
            "entregas": [],
            "documentos": [],
            "guia_ultima": "",
            "fatura_ultima": "",
            "fatura_caminho_ultima": "",
            "data_doc_ultima": "",
            "data_ultima_entrega": "",
        }
        notes.append(note)
        self.repository.replace(notes, expected=original)
        return deepcopy(note)


    def remove(self, numero: str) -> None:
        original = self.repository.rows()
        notes = deepcopy(original)
        numero = str(numero or "").strip()
        before = len(notes)
        notes = [row for row in notes if str(row.get("numero", "") or "").strip() != numero]
        if len(notes) == before:
            raise ValueError("Nota de Encomenda não encontrada.")
        self.repository.replace(notes, expected=original)


    def approve(self, numero: str) -> dict[str, Any]:
        original = self.repository.rows()
        notes = deepcopy(original)
        numero = str(numero or "").strip()
        note = next((row for row in notes if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda não encontrada.")
        if not list(note.get("linhas", []) or []):
            raise ValueError("A nota não tem linhas.")
        note_kind = self.rules.kind(note)
        note["estado"] = "Aprovada" if note_kind == "purchase_note" else "Cotacao aprovada"
        note["data_aprovacao"] = self.rules.now_iso()
        note["_draft"] = False
        self.repository.replace(notes, expected=original)
        return deepcopy(note)


    def mark_sent(self, numero: str) -> dict[str, Any]:
        original = self.repository.rows()
        notes = deepcopy(original)
        numero = str(numero or "").strip()
        note = next((row for row in notes if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda não encontrada.")
        note["estado"] = "Enviada"
        note["data_envio"] = self.rules.now_iso()
        note["_draft"] = False
        self.repository.replace(notes, expected=original)
        return deepcopy(note)


    def generate_orders(self, numero: str) -> list[dict[str, Any]]:
        original = self.repository.rows()
        notes = deepcopy(original)
        number = str(numero or "").strip()
        note = next((row for row in notes if str(row.get("numero", "") or "").strip() == number), None)
        if note is None:
            raise ValueError("Nota de Encomenda não encontrada.")
        if note.get("ne_geradas") or note.get("estado") == "Convertida":
            raise ValueError("A cotacao ja foi convertida em encomendas a fornecedores.")
        lines = list(note.get("linhas", []) or [])
        if not lines:
            raise ValueError("A nota não tem linhas.")
        groups: dict[str, dict[str, Any]] = {}
        missing: list[str] = []
        for line in lines:
            target_supplier = str(line.get("fornecedor_linha", "") or note.get("fornecedor", "") or "").strip()
            supplier_id, supplier_text, supplier_contact = self.rules.supplier(target_supplier)
            if not supplier_id or not supplier_text:
                missing.append(str(line.get("ref", "") or "").strip() or str(line.get("descricao", "") or "").strip())
                continue
            key = supplier_id or supplier_text
            if key not in groups:
                groups[key] = {
                    "fornecedor_id": supplier_id,
                    "fornecedor": supplier_text,
                    "contacto": supplier_contact,
                    "linhas": [],
                }
            new_line = dict(line)
            new_line["fornecedor_linha"] = supplier_text
            new_line["entregue"] = False
            new_line["qtd_entregue"] = 0.0
            new_line["_stock_in"] = False
            for transient_key in ("guia_entrega", "fatura_entrega", "data_doc_entrega", "data_entrega_real", "obs_entrega", "entregas_linha"):
                new_line.pop(transient_key, None)
            groups[key]["linhas"].append(new_line)
        if missing:
            raise ValueError("Existem linhas sem fornecedor adjudicado: " + ", ".join(sorted(set(item for item in missing if item))))
        if not groups:
            raise ValueError("Nao existem fornecedores adjudicados para gerar NEs.")
        created: list[dict[str, Any]] = []
        for group in groups.values():
            new_number = str(self.repository.allocate_number())
            new_note = {
                "numero": new_number,
                "fornecedor": group.get("fornecedor", ""),
                "fornecedor_id": group.get("fornecedor_id", ""),
                "contacto": group.get("contacto", ""),
                "created_at": self.rules.now_iso(),
                "referencias_orcamento": str(note.get("referencias_orcamento", "") or "").strip(),
                "data_entrega": str(note.get("data_entrega", "") or "").strip(),
                "obs": f"Gerada de {note.get('numero', '')}".strip(),
                "local_descarga": str(note.get("local_descarga", "") or "").strip(),
                "meio_transporte": str(note.get("meio_transporte", "") or "").strip(),
                "linhas": list(group.get("linhas", []) or []),
                "estado": "Aprovada",
                "_draft": False,
                "oculta": False,
                "origem_cotacao": str(note.get("numero", "") or "").strip(),
                "ne_geradas": [],
                "entregas": [],
                "documentos": [],
                "guia_ultima": "",
                "fatura_ultima": "",
                "fatura_caminho_ultima": "",
                "data_doc_ultima": "",
                "data_ultima_entrega": "",
            }
            self.rules.recalculate(new_note)
            notes.append(new_note)
            created.append({"numero": new_number, "fornecedor": new_note["fornecedor"], "total": new_note["total"]})
        note["estado"] = "Convertida"
        note["oculta"] = True
        note["_draft"] = False
        note["ne_geradas"] = [row["numero"] for row in created]
        self.repository.replace(notes, expected=original)
        return created
