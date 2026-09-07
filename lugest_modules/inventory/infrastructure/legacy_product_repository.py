"""Read-only product collections, detached from the historical snapshot."""
from copy import deepcopy
from typing import Callable


class LegacyProductReadRepository:
    def __init__(self, *, products: Callable[[], list[dict]], movements: Callable[[], list[dict]]):
        self.read_products = products
        self.read_movements = movements

    def products(self) -> list[dict]:
        return deepcopy(list(self.read_products() or []))

    def movements(self) -> list[dict]:
        return deepcopy(list(self.read_movements() or []))
