"""Client use cases with an explicit repository; no widgets or legacy globals."""
from __future__ import annotations

from typing import Any, Callable, Protocol

from lugest_core.search import search_normalize


CLIENT_FIELDS = (
    "codigo", "nome", "nif", "morada", "latitude", "longitude", "contacto",
    "email", "observacoes", "prazo_entrega", "cond_pagamento", "obs_tecnicas",
)


class ClientRepository(Protocol):
    def list_clients(self) -> list[dict[str, Any]]: ...
    def next_code(self) -> str: ...
    def save(self, row: dict[str, Any]) -> dict[str, Any]: ...
    def reference_kinds(self, code: str) -> set[str]: ...
    def remove(self, code: str) -> None: ...


class ClientService:
    def __init__(self, repository: ClientRepository, *,
                 normalize_search: Callable[[Any], str] = search_normalize) -> None:
        self.repository = repository
        self.normalize_search = normalize_search

    def rows(self, filter_text: str = "") -> list[dict[str, str]]:
        terms = self.normalize_search(filter_text).split()
        rows = []
        for raw in self.repository.list_clients():
            row = {key: str(raw.get(key, "") or "").strip() for key in CLIENT_FIELDS}
            haystack = self.normalize_search(" ".join(row.values()))
            if all(term in haystack for term in terms):
                rows.append(row)
        return sorted(rows, key=lambda row: (row["codigo"], row["nome"]))

    def order_options(self) -> list[dict[str, str]]:
        rows = []
        for raw in self.repository.list_clients():
            code, name = (str(raw.get(key, "") or "").strip() for key in ("codigo", "nome"))
            if code:
                rows.append({"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")})
        return sorted(rows, key=lambda row: row["codigo"])

    def next_code(self) -> str:
        return self.repository.next_code()

    def save(self, payload: dict[str, Any]) -> dict[str, Any]:
        row = {key: str(payload.get(key, "") or "").strip() for key in CLIENT_FIELDS}
        if not row["nome"]:
            raise ValueError("Nome do cliente obrigatorio.")
        if not row["codigo"]:
            row["codigo"] = self.repository.next_code()
        return self.repository.save(row)

    def remove(self, code: str) -> None:
        code = str(code or "").strip()
        if not code:
            raise ValueError("Cliente inválido.")
        references = self.repository.reference_kinds(code)
        if "orders" in references:
            raise ValueError("Nao e possivel remover um cliente usado em encomendas.")
        if "quotes" in references:
            raise ValueError("Nao e possivel remover um cliente usado em orcamentos.")
        self.repository.remove(code)
