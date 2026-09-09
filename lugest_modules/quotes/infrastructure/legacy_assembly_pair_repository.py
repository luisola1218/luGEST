"""Publish both prepared catalogs with one blocking dataset save."""
from copy import deepcopy


class LegacyAssemblyPairRepository:
    def __init__(self, get_data, save_dataset):
        self.get_data = get_data
        self.save_dataset = save_dataset

    def replace(self, template, live):
        prepared = {"conjuntos_modelo": template, "conjuntos": live}
        data = self.get_data()
        for key, change in prepared.items():
            if list(data.get(key, []) or []) != change.expected:
                raise ValueError("Os conjuntos foram alterados. Atualize e tente novamente.")
        previous = {key: deepcopy(data[key]) for key in prepared if key in data}
        for key, change in prepared.items():
            data[key] = deepcopy(change.models)
        try:
            self.save_dataset(force=True, blocking=True)
        except Exception:
            current = self.get_data()
            for key in prepared:
                if key in previous:
                    current[key] = previous[key]
                else:
                    current.pop(key, None)
            raise
