"""Compatibility repository with explicit callbacks to existing persistence.

Synchronous upserts publish to the cache only after success. Dataset writes
restore the client collection on a synchronous error. Asynchronous acceptance
does not guarantee durability; the existing save worker still reports failures.
"""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Callable


class LegacyClientRepository:
    def __init__(self, *, get_data: Callable[[], dict], get_baseline: Callable[[], dict | None],
                 save_dataset: Callable[..., None], next_code: Callable[[dict], str],
                 normalize_quote_client: Callable[[Any], dict],
                 upsert: Callable[[dict], None] | None = None, asynchronous: bool = False) -> None:
        self.get_data = get_data
        self.get_baseline = get_baseline
        self.save_dataset = save_dataset
        self.generate_code = next_code
        self.normalize_quote_client = normalize_quote_client
        self.upsert = upsert
        self.asynchronous = asynchronous

    def list_clients(self) -> list[dict]:
        return deepcopy(list(self.get_data().get("clientes", []) or []))

    def next_code(self) -> str:
        return str(self.generate_code(self.get_data()))

    @staticmethod
    def _replace(rows: list[dict], target: dict) -> list[dict]:
        code = target["codigo"]
        for index, item in enumerate(rows):
            if str(item.get("codigo", "") or "").strip() == code:
                return [*rows[:index], target, *rows[index + 1:]]
        return [*rows, target]

    def save(self, row: dict) -> dict:
        data = self.get_data()
        previous = data.get("clientes", [])
        existing = next((item for item in previous if str(item.get("codigo", "") or "").strip() == row["codigo"]), {})
        target = {**deepcopy(existing), **deepcopy(row)}
        if not self.asynchronous and self.upsert is not None:
            self.upsert(deepcopy(target))
            data["clientes"] = self._replace(previous, target)
            baseline = self.get_baseline()
            if isinstance(baseline, dict):
                baseline["clientes"] = self._replace(baseline.get("clientes", []), deepcopy(target))
        else:
            data["clientes"] = self._replace(previous, target)
            try:
                self.save_dataset(force=not self.asynchronous)
            except Exception:
                data["clientes"] = previous
                raise
        return deepcopy(target)

    def reference_kinds(self, code: str) -> set[str]:
        data = self.get_data()
        references = set()
        if any(str(row.get("cliente", "") or "").strip() == code for row in data.get("encomendas", []) or []):
            references.add("orders")
        if any(str(self.normalize_quote_client(row.get("cliente", {})).get("codigo", "") or "").strip() == code
               for row in data.get("orcamentos", []) or []):
            references.add("quotes")
        return references

    def remove(self, code: str) -> None:
        data = self.get_data()
        previous = data.get("clientes", [])
        remaining = [row for row in previous if str(row.get("codigo", "") or "").strip() != code]
        if len(previous) == len(remaining):
            raise ValueError("Cliente não encontrado.")
        data["clientes"] = remaining
        try:
            self.save_dataset(force=True)
        except Exception:
            data["clientes"] = previous
            raise
