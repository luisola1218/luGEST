"""Compatibility write adapter; restores cache changes on an immediate failure."""
from copy import deepcopy
from typing import Callable


class LegacyStockIssueRepository:
    def __init__(self, *, get_data: Callable[[], dict], save_dataset: Callable[..., None],
                 add_movement: Callable[..., None], parse_float: Callable[..., float]):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.add_movement = add_movement
        self.parse_float = parse_float

    @staticmethod
    def _find(data, code):
        return next((row for row in data.get('produtos', [])
                     if str(row.get('codigo', '') or '').strip() == code), None)

    def product(self, code):
        return deepcopy(self._find(self.get_data(), code))

    def issue(self, code, *, expected_quantity, quantity_after, updated_at, movement):
        data = self.get_data()
        row = self._find(data, code)
        if row is None:
            raise ValueError('Produto não encontrado.')
        if abs(self.parse_float(row.get('qty', 0), 0) - expected_quantity) > 1e-9:
            raise ValueError('O stock foi alterado. Atualiza o produto e tenta novamente.')
        previous = {key: deepcopy(row[key]) for key in ('qty', 'atualizado_em') if key in row}
        had_movements = 'produtos_mov' in data
        previous_movements = deepcopy(data.get('produtos_mov', []))
        try:
            row['qty'] = quantity_after
            row['atualizado_em'] = updated_at
            self.add_movement(data, **deepcopy(movement))
            self.save_dataset(force=True)
        except Exception:
            current_data = self.get_data()
            current = self._find(current_data, code)
            if current is not None:
                for key in ('qty', 'atualizado_em'):
                    if key in previous:
                        current[key] = previous[key]
                    else:
                        current.pop(key, None)
            if had_movements:
                current_data['produtos_mov'] = previous_movements
            else:
                current_data.pop('produtos_mov', None)
            raise
