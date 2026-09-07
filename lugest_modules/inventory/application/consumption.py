"""Stock issue use case: validate the complete command before any mutation."""
from __future__ import annotations

from math import isfinite
from typing import Any, Callable, Protocol


class StockIssueRepository(Protocol):
    def product(self, code: str) -> dict | None: ...
    def issue(self, code: str, *, expected_quantity: float, quantity_after: float,
              updated_at: str, movement: dict) -> None: ...


class StockIssueService:
    def __init__(self, repository: StockIssueRepository, *, parse_float: Callable[..., float],
                 unit_price: Callable[[dict], float], now: Callable[[], str]):
        self.repository = repository
        self.parse_float = parse_float
        self.unit_price = unit_price
        self.now = now

    def consume(self, code: str, quantity: Any, *, actor: str, observation: str = '',
                target_operator: str = '', issue_mode: str = 'stock') -> dict:
        code = str(code or '').strip()
        quantity = self.parse_float(quantity, 0)
        if not isfinite(quantity) or quantity <= 0:
            raise ValueError('Quantidade invalida.')
        product = self.repository.product(code)
        if product is None:
            raise ValueError('Produto não encontrado.')
        before = self.parse_float(product.get('qty', 0), 0)
        if quantity > before + 1e-9:
            raise ValueError('Quantidade superior ao stock disponivel.')
        actor = str(actor or 'Sistema')
        operator = str(target_operator or '').strip()
        issue_mode = str(issue_mode or 'stock').strip().lower()
        movement_type = 'BAIXA'
        note = str(observation or '').strip() or 'Baixa manual no desktop Qt'
        if issue_mode == 'operator':
            if not operator:
                raise ValueError('Seleciona o operador que recebe o material.')
            movement_type = 'ENTREGA_OPERADOR'
            unit_price = round(self.parse_float(self.unit_price(product), 0), 4)
            total = round(unit_price * quantity, 2)
            note = str(observation or '').strip() or 'Entrega a operador'
            note = f'{note} |meta|actor={actor} |meta|valor_unit={unit_price:.4f} |meta|valor_total={total:.2f}'
        after = max(0.0, before - quantity)
        movement = {
            'tipo': movement_type, 'operador': operator or actor, 'codigo': code,
            'descricao': str(product.get('descricao', '') or '').strip(),
            'qtd': quantity, 'antes': before, 'depois': after, 'obs': note,
            'origem': 'OPERADOR' if movement_type == 'ENTREGA_OPERADOR' else 'PRODUTOS',
            'ref_doc': code,
        }
        self.repository.issue(code, expected_quantity=before, quantity_after=after,
                              updated_at=self.now(), movement=movement)
        return {'codigo': code, 'quantidade': quantity, 'antes': before, 'depois': after}
