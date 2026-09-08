"""Quote aggregate persistence adapted to the historical runtime.

Registry callbacks retain their existing SQL behavior. Recovery below restores
the local cache on immediate failure; it is not a cross-connection transaction.
"""
from copy import deepcopy
from typing import Callable


class LegacyQuoteWriteRepository:
    FIELDS = ('orcamentos', 'orc_seq', 'orc_refs', 'peca_hist')

    def __init__(self, *, get_data: Callable[[], dict], get_baseline: Callable[[], dict | None],
                 save_dataset: Callable[..., None], next_number: Callable[[], str],
                 next_reference: Callable[[str, list[str]], str], peek_number: Callable[[], str],
                 sync_registry: Callable[[dict], None], upsert: Callable[[dict, dict], None] | None,
                 delete_studies: Callable[[str], None]):
        self.get_data = get_data
        self.get_baseline = get_baseline
        self.save_dataset = save_dataset
        self.allocate_number = next_number
        self.allocate_reference = next_reference
        self.peek_number = peek_number
        self.sync_registry = sync_registry
        self.upsert = upsert
        self.delete_studies = delete_studies

    @staticmethod
    def _find(data, number):
        return next((row for row in data.get('orcamentos', [])
                     if str(row.get('numero', '') or '').strip() == number), None)

    def get(self, number):
        return deepcopy(self._find(self.get_data(), number))

    def next_number(self):
        return str(self.allocate_number())

    def next_reference(self, client, reserved):
        return str(self.allocate_reference(client, reserved))

    def _restore(self, before, fields):
        current = self.get_data()
        for key in fields:
            if key in before:
                current[key] = before[key]
            else:
                current.pop(key, None)

    def save(self, quote, *, direct=True):
        data = self.get_data()
        number = quote['numero']
        previous = {key: deepcopy(data[key]) for key in self.FIELDS if key in data}
        target = self._find(data, number)
        try:
            if target is None:
                target = deepcopy(quote)
                data.setdefault('orcamentos', []).append(target)
                if number == self.peek_number():
                    try:
                        data['orc_seq'] = max(int(data.get('orc_seq', 1) or 1), int(number.rsplit('-', 1)[-1]) + 1)
                    except (TypeError, ValueError):
                        pass
            else:
                target.update(deepcopy(quote))
            self.sync_registry(target)
            data = self.get_data()
            target = self._find(data, number)
            if target is None:
                raise ValueError('Orçamento não encontrado.')
            if direct and callable(self.upsert):
                self.upsert(data, target)
                baseline = self.get_baseline()
                if isinstance(baseline, dict):
                    existing = self._find(baseline, number)
                    if existing is None:
                        baseline.setdefault('orcamentos', []).append(deepcopy(target))
                    else:
                        existing.update(deepcopy(target))
            else:
                self.save_dataset(force=True)
        except Exception:
            self._restore(previous, self.FIELDS)
            raise

    def remove(self, number):
        data = self.get_data()
        rows = list(data.get('orcamentos', []) or [])
        kept = [row for row in rows if str(row.get('numero', '') or '').strip() != number]
        if len(kept) == len(rows):
            raise ValueError('Orçamento não encontrado.')
        previous = {'orcamentos': deepcopy(rows)}
        try:
            data['orcamentos'] = kept
            self.save_dataset(force=True)
        except Exception:
            self._restore(previous, ('orcamentos',))
            raise
        try:
            self.delete_studies(number)
        except Exception:
            pass
