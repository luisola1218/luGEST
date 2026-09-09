"""Quality document metadata and audit operations on detached records."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class DocumentRepository(Protocol):
    def rows(self) -> list[dict[str, Any]]: ...
    def replace(self, rows: list[dict[str, Any]], *, expected: list[dict[str, Any]], event: dict[str, Any]) -> None: ...

@dataclass(frozen=True)
class DocumentRules:
    next_id: Callable
    store_file: Callable
    now_iso: Callable
    actor: Callable

class QualityDocuments:
    def __init__(self, repository: DocumentRepository, rules: DocumentRules):
        self.repository = repository
        self.rules = rules

    def rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in self.repository.rows():
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


    def save(self, payload: dict[str, Any]) -> dict[str, Any]:
        original = self.repository.rows()
        rows = deepcopy(original)
        doc_id = str(payload.get("id", "") or "").strip()
        existing = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == doc_id), None) if doc_id else None
        before = deepcopy(existing) if isinstance(existing, dict) else None
        if not doc_id:
            doc_id = self.rules.next_id(rows, "DOC")
        titulo = str(payload.get("titulo", "") or "").strip()
        if not titulo:
            raise ValueError("Titulo do documento obrigatorio.")
        source_path = str(payload.get("caminho", "") or "").strip()
        stored_path = source_path
        if source_path:
            stored_path = self.rules.store_file(source_path, titulo or doc_id)
        now = str(self.rules.now_iso())
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
            "created_by": str((existing or {}).get("created_by", "") or self.rules.actor()),
            "updated_by": self.rules.actor(),
        }
        if existing is None:
            rows.append(row)
        else:
            existing.update(row)
            row = existing
        event = dict(action="Documento qualidade guardado", entity_type="Documento", entity_id=doc_id, summary=titulo, before=before, after=row)
        self.repository.replace(rows, expected=original, event=event)
        return deepcopy(row)


    def remove(self, doc_id: str) -> None:
        original = self.repository.rows()
        value = str(doc_id or "").strip()
        rows = deepcopy(original)
        before = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == value), None)
        updated = [row for row in rows if not (isinstance(row, dict) and str(row.get("id", "") or "").strip() == value)]
        if before is None:
            raise ValueError("Documento nao encontrado.")
        event = dict(action="Documento qualidade removido", entity_type="Documento", entity_id=value, summary=str(before.get("titulo", "") or ""), before=before)
        self.repository.replace(updated, expected=original, event=event)
