"""Product writes against the current snapshot, with immediate failure recovery.

The legacy save callback defines durability. This adapter does not claim a
transaction across clients or undo a write already committed by that callback.
"""
from copy import deepcopy
from typing import Callable


class LegacyProductWriteRepository:
    def __init__(self, *, get_data: Callable[[], dict], save_dataset: Callable[..., None],
                 add_movement: Callable[..., None], ensure_sequence: Callable[[dict, str], None]):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.add_movement = add_movement
        self.ensure_sequence = ensure_sequence

    @staticmethod
    def _find(data, code):
        return next((row for row in data.get('produtos', [])
                     if str(row.get('codigo', '') or '').strip() == code), None)

    def product(self, code):
        return deepcopy(self._find(self.get_data(), code))

    def _restore(self, before, fields):
        current = self.get_data()
        for key in fields:
            if key in before:
                current[key] = before[key]
            else:
                current.pop(key, None)

    def save(self, product, *, expected, movement):
        data = self.get_data()
        code = product['codigo']
        current = self._find(data, code)
        if current != expected:
            raise ValueError('O produto foi alterado. Atualiza o registo e tenta novamente.')
        fields = ('produtos', 'produtos_mov', 'seq')
        before = {key: deepcopy(data[key]) for key in fields if key in data}
        try:
            if current is None:
                data.setdefault('produtos', []).append(deepcopy(product))
            else:
                current.update(deepcopy(product))
            if movement is not None:
                self.add_movement(data, **deepcopy(movement))
            self.ensure_sequence(data, code)
            self.save_dataset(force=True)
        except Exception:
            self._restore(before, fields)
            raise

    def remove(self, codes):
        data = self.get_data()
        rows = list(data.get('produtos', []) or [])
        kept = [row for row in rows if str(row.get('codigo', '') or '').strip() not in codes]
        count = len(rows) - len(kept)
        if not count:
            raise ValueError('Os produtos selecionados já não existem.')
        before = {'produtos': deepcopy(rows)}
        try:
            data['produtos'] = kept
            self.save_dataset(force=True)
        except Exception:
            self._restore(before, ('produtos',))
            raise
        return count
