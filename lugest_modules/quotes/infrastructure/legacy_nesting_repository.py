"""Adapter for the historical dataset and its optional SQL study mirror.

The dataset remains authoritative for writes. The SQL mirror is best effort,
as before; it is not an atomic second database write. An immediate dataset
failure restores the affected fields in the current cache.
"""
from __future__ import annotations

from copy import deepcopy
from typing import Callable


class LegacyNestingStudyRepository:
    FIELDS = ('nesting_studies', 'latest_nesting_bridge', 'latest_nesting_group_key',
              'latest_nesting_updated_at')

    def __init__(self, *, get_data: Callable[[], dict], save_dataset: Callable[..., None],
                 load_remote: Callable[[str], dict], save_remote: Callable[..., None]):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.load_remote = load_remote
        self.save_remote = save_remote

    def _record(self, number: str) -> dict | None:
        return next((row for row in self.get_data().get('orcamentos', [])
                     if str(row.get('numero', '') or '').strip() == number), None)

    def quote(self, number: str) -> dict | None:
        row = self._record(number)
        return deepcopy(row) if row is not None else None

    def remote_studies(self, number: str) -> dict:
        return deepcopy(self.load_remote(number))

    def save(self, number: str, study: dict) -> None:
        row = self._record(number)
        if row is None:
            raise ValueError('Orçamento não encontrado.')
        previous = {key: deepcopy(row[key]) for key in self.FIELDS if key in row}
        key = study['group_key']
        row.setdefault('nesting_studies', {})[key] = deepcopy(study)
        row['latest_nesting_bridge'] = deepcopy(study.get('quote_bridge', {}) or {})
        row['latest_nesting_group_key'] = key
        row['latest_nesting_updated_at'] = study['updated_at']
        try:
            self.save_dataset(force=True)
        except Exception:
            current = self._record(number)
            if current is not None:
                for field in self.FIELDS:
                    if field in previous:
                        current[field] = previous[field]
                    else:
                        current.pop(field, None)
            raise
        try:
            self.save_remote(number, key, study.get('group_label', ''), deepcopy(study))
        except Exception:
            # Compatibility: the dataset already contains the saved study.
            pass
