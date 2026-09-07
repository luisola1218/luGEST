"""Nesting study use cases with explicit persistence and clock dependencies."""
from __future__ import annotations

from copy import deepcopy
import json
from typing import Any, Callable, Protocol


class NestingStudyRepository(Protocol):
    def quote(self, number: str) -> dict | None: ...
    def remote_studies(self, number: str) -> dict: ...
    def save(self, number: str, study: dict) -> None: ...


def json_clone(value: Any) -> Any:
    return json.loads(json.dumps(value, ensure_ascii=False, default=str))


class NestingStudyService:
    def __init__(self, repository: NestingStudyRepository, *, now: Callable[[], str]):
        self.repository = repository
        self.now = now

    def studies(self, number: str) -> dict:
        number = str(number or '').strip()
        quote = self.repository.quote(number) or {}
        local = {str(key): json_clone(value)
                 for key, value in dict(quote.get('nesting_studies', {}) or {}).items()
                 if str(key).strip()}
        for key, remote in self.repository.remote_studies(number).items():
            previous = dict(local.get(key, {}) or {})
            remote_updated = str(dict(remote or {}).get('updated_at', '') or '').strip()
            local_updated = str(previous.get('updated_at', '') or '').strip()
            if not previous or remote_updated >= local_updated:
                local[key] = json_clone(remote)
        return local

    def save(self, number: str, payload: dict) -> dict:
        number = str(number or '').strip()
        quote = self.repository.quote(number)
        if quote is None:
            raise ValueError('Guarda primeiro o orçamento para associar o estudo de nesting.')
        clean = dict(json_clone(payload) or {})
        key = str(clean.get('group_key', '') or '').strip()
        if not key:
            raise ValueError('Grupo de nesting inválido.')
        previous = dict(dict(quote.get('nesting_studies', {}) or {}).get(key, {}) or {})
        clean['quote_number'] = number
        clean['group_key'] = key
        clean['group_label'] = str(clean.get('group_label', previous.get('group_label', '')) or '').strip()
        clean['created_at'] = str(previous.get('created_at', '') or clean.get('created_at', '') or self.now()).strip()
        clean['updated_at'] = self.now()
        self.repository.save(number, deepcopy(clean))
        return deepcopy(clean)
