"""Nonconformity lifecycle and read models with explicit audit persistence."""
import copy
import re
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class NonconformityRepository(Protocol):
    def rows(self) -> list[dict[str, Any]]: ...
    def replace(self, rows: list[dict[str, Any]], *, expected: list[dict[str, Any]], event: dict[str, Any] | None = None) -> None: ...

@dataclass(frozen=True)
class NonconformityRules:
    parse_float: Callable
    now_iso: Callable
    actor: Callable
    next_id: Callable
    link_label: Callable

class Nonconformities:
    def __init__(self, repository: NonconformityRepository, rules: NonconformityRules):
        self.repository = repository
        self.rules = rules

    def reference_key(self, value: Any, fallback: Any = "") -> str:
        raw = str(value or fallback or "").strip()
        match = re.search(r"\bNE-\d{4}-\d{4}\b", raw, flags=re.IGNORECASE)
        if match:
            return match.group(0).upper()
        return re.sub(r"\s+", " ", raw).casefold()


    def key(self, payload: dict[str, Any]) -> tuple[str, str, str, str]:
        origem_raw = str(payload.get("origem", "") or "").strip()
        origem = re.sub(r"\s+", " ", origem_raw).casefold()
        if "rece" in origem and "fornecedor" in origem:
            origem = "rececao fornecedor"
        referencia = self.reference_key(payload.get("referencia", ""), payload.get("ne_numero", ""))
        entidade_tipo = str(payload.get("entidade_tipo", "") or payload.get("linked_entity_type", "") or "").strip()
        entidade_id = str(payload.get("entidade_id", "") or payload.get("linked_entity_id", "") or "").strip()
        if not entidade_tipo and str(payload.get("material_id", "") or "").strip():
            entidade_tipo = "Material"
            entidade_id = str(payload.get("material_id", "") or "").strip()
        if not entidade_id:
            entidade_id = str(payload.get("material_id", "") or payload.get("produto_codigo", "") or payload.get("fornecedor_id", "") or payload.get("fornecedor_nome", "") or "").strip()
        return (origem, referencia, entidade_tipo.casefold(), entidade_id.casefold())


    def is_open(self, row: dict[str, Any]) -> bool:
        return str(row.get("estado", "") or "Aberta").strip().casefold() == "aberta"


    def find_open(self, payload: dict[str, Any], *, exclude_id: str = "") -> dict[str, Any] | None:
        key = self.key(payload)
        exclude = str(exclude_id or "").strip()
        for row in self.repository.rows():
            if not isinstance(row, dict) or not self.is_open(row):
                continue
            if exclude and str(row.get("id", "") or "").strip() == exclude:
                continue
            if self.key(row) == key:
                return row
        return None


    def quantity(self, row: dict[str, Any] | None, field: str) -> float:
        if not isinstance(row, dict):
            return 0.0
        for key in (field, field.replace("qtd_", "quality_"), field.replace("qtd_", "")):
            if key in row:
                value = self.rules.parse_float(row.get(key, 0), 0)
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
            return self.rules.parse_float(match.group(1), 0)
        return 0.0


    def normalize_duplicates(self) -> None:
        original = self.repository.rows()
        rows = copy.deepcopy(original)
        first_by_key: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        changed = False
        for row in rows:
            if not isinstance(row, dict) or not self.is_open(row):
                continue
            key = self.key(row)
            if not all(key):
                continue
            keeper = first_by_key.get(key)
            if keeper is None:
                first_by_key[key] = row
                continue
            row["estado"] = "Cancelada"
            row["updated_at"] = str(self.rules.now_iso())
            row["updated_by"] = self.rules.actor()
            row["acao"] = (
                str(row.get("acao", "") or "").strip()
                + f"\nCancelada automaticamente: NC duplicada de {str(keeper.get('id', '') or '').strip()}."
            ).strip()
            changed = True
        if changed:
            self.repository.replace(rows, expected=original)


    def rows(self, filter_text: str = "", state_filter: str = "Ativas") -> list[dict[str, Any]]:
        self.normalize_duplicates()
        query = str(filter_text or "").strip().lower()
        state = str(state_filter or "Ativas").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in self.repository.rows():
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
                "qtd_recebida": round(self.quantity(row, "qtd_recebida"), 4),
                "qtd_aprovada": round(self.quantity(row, "qtd_aprovada"), 4),
                "qtd_rejeitada": round(self.quantity(row, "qtd_rejeitada"), 4),
                "qtd_pendente": round(self.quantity(row, "qtd_pendente"), 4),
                "created_at": str(row.get("created_at", "") or "").strip(),
                "closed_at": str(row.get("closed_at", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in emitted.values()):
                continue
            rows.append(emitted)
        rows.sort(key=lambda item: (str(item.get("estado", "")) == "Fechada", str(item.get("prazo", "") or "9999"), str(item.get("id", ""))), reverse=False)
        return rows


    def save(self, payload: dict[str, Any]) -> dict[str, Any]:
        original = self.repository.rows()
        rows = copy.deepcopy(original)
        nc_id = str(payload.get("id", "") or "").strip()
        existing = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == nc_id), None) if nc_id else None
        before = copy.deepcopy(existing) if isinstance(existing, dict) else None
        if not nc_id:
            nc_id = self.rules.next_id(rows, "NC")
        now = str(self.rules.now_iso())
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
            "qtd_recebida": round(self.rules.parse_float(payload.get("qtd_recebida", (existing or {}).get("qtd_recebida", 0)), 0), 4),
            "qtd_aprovada": round(self.rules.parse_float(payload.get("qtd_aprovada", (existing or {}).get("qtd_aprovada", 0)), 0), 4),
            "qtd_rejeitada": round(self.rules.parse_float(payload.get("qtd_rejeitada", (existing or {}).get("qtd_rejeitada", 0)), 0), 4),
            "qtd_pendente": round(self.rules.parse_float(payload.get("qtd_pendente", (existing or {}).get("qtd_pendente", 0)), 0), 4),
            "created_at": str((existing or {}).get("created_at", "") or now),
            "updated_at": now,
            "created_by": str((existing or {}).get("created_by", "") or self.rules.actor()),
            "updated_by": self.rules.actor(),
            "closed_at": str((existing or {}).get("closed_at", "") or "").strip(),
        }
        row["entidade_label"] = str(payload.get("entidade_label", "") or "").strip() or self.rules.link_label(
            row["entidade_tipo"], row["entidade_id"]
        )
        if not row["referencia"] and row["entidade_id"]:
            row["referencia"] = row["entidade_id"]
        if self.is_open(row):
            duplicate = self.find_open(row, exclude_id=nc_id)
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
        event = dict(
            action="NC guardada",
            entity_type="Nao conformidade",
            entity_id=nc_id,
            summary=f"{row.get('tipo', '')} | {row.get('estado', '')} | {row.get('referencia', '')}",
            before=before,
            after=row,
        )
        self.repository.replace(rows, expected=original, event=event)
        return copy.deepcopy(row)


    def close(self, nc_id: str, eficacia: str = "") -> dict[str, Any]:
        original = self.repository.rows()
        rows = copy.deepcopy(original)
        target = next((row for row in rows if isinstance(row, dict) and str(row.get("id", "") or "").strip() == str(nc_id or "").strip()), None)
        if target is None:
            raise ValueError("Nao conformidade nao encontrada.")
        before = copy.deepcopy(target)
        target["estado"] = "Fechada"
        target["closed_at"] = str(self.rules.now_iso())
        target["closed_by"] = self.rules.actor()
        if str(eficacia or "").strip():
            target["eficacia"] = str(eficacia or "").strip()
        event = dict( action="NC fechada", entity_type="Nao conformidade", entity_id=str(nc_id), summary=str(target.get("eficacia", "") or ""), before=before, after=target)
        self.repository.replace(rows, expected=original, event=event)
        return copy.deepcopy(target)



    def remove(self, nc_id: str) -> None:
        original = self.repository.rows()
        value = str(nc_id or "").strip()
        before = next((row for row in original if isinstance(row, dict) and str(row.get("id", "") or "").strip() == value), None)
        if before is None:
            raise ValueError("Nao conformidade nao encontrada.")
        rows = [row for row in original if not (isinstance(row, dict) and str(row.get("id", "") or "").strip() == value)]
        event = dict(action="NC removida", entity_type="Nao conformidade", entity_id=value,
                     summary=str(before.get("descricao", "") or ""), before=before)
        self.repository.replace(rows, expected=original, event=event)
