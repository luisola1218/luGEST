"""Create, edit and remove products without owning the application snapshot."""
from copy import deepcopy
from typing import Callable, Protocol


class ProductWriteRepository(Protocol):
    def product(self, code: str) -> dict | None: ...
    def save(self, product: dict, *, expected: dict | None, movement: dict | None) -> None: ...
    def remove(self, codes: set[str]) -> int: ...


class ProductCommands:
    def __init__(self, repository: ProductWriteRepository, *, normalize: Callable[[dict], dict],
                 parse_float: Callable[..., float], format_number: Callable[..., str],
                 now: Callable[[], str]):
        self.repository = repository
        self.normalize = normalize
        self.parse_float = parse_float
        self.format_number = format_number
        self.now = now

    def save(self, payload: dict, *, actor: str) -> str:
        normalized = self.normalize(deepcopy(payload))
        code = str(normalized.get('codigo', '') or '').strip()
        previous = self.repository.product(code)
        target = {**deepcopy(previous or {}), **normalized, 'atualizado_em': self.now()}
        before = self.parse_float((previous or {}).get('qty', 0), 0)
        after = self.parse_float(target.get('qty', 0), 0)
        movement = None
        if previous is None and after > 1e-9:
            movement = {'tipo': 'ENTRADA_INICIAL', 'qtd': after,
                        'obs': 'Stock inicial no registo do produto'}
        elif previous is not None and abs(after - before) > 1e-9:
            movement = {'tipo': 'AJUSTE_STOCK', 'qtd': abs(after - before),
                        'obs': f'Ajuste manual no cadastro ({self.format_number(after - before)})'}
        if movement is not None:
            movement.update(operador=actor, codigo=code,
                            descricao=str(target.get('descricao', '') or '').strip(),
                            antes=before, depois=after, origem='PRODUTOS', ref_doc=code)
        self.repository.save(target, expected=previous, movement=movement)
        return code

    def remove(self, codes: list[str]) -> int:
        codes = {str(value or '').strip() for value in codes if str(value or '').strip()}
        if not codes:
            raise ValueError('Seleciona pelo menos um produto.')
        return self.repository.remove(codes)
