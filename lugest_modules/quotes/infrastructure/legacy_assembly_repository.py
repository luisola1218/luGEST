"""Snapshot adapter; recovery covers immediate errors, not asynchronous commits."""
from copy import deepcopy
from typing import Callable


class LegacyAssemblyRepository:
    def __init__(self, get_data: Callable, save_dataset: Callable, collection: str = "conjuntos"):
        if collection not in {"conjuntos", "conjuntos_modelo"}:
            raise ValueError("Unknown assembly catalog")
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.collection = collection

    def models(self):
        return deepcopy(list(self.get_data().get(self.collection, []) or []))

    def replace(self, models, *, expected):
        data = self.get_data()
        if list(data.get(self.collection, []) or []) != expected:
            raise ValueError("Os conjuntos foram alterados. Atualize e tente novamente.")
        present = self.collection in data
        previous = deepcopy(data.get(self.collection))
        data[self.collection] = deepcopy(models)
        try:
            self.save_dataset(force=True)
        except Exception:
            current = self.get_data()
            if present:
                current[self.collection] = previous
            else:
                current.pop(self.collection, None)
            raise
