"""Detached tariff snapshots and recovery from immediate persistence failures."""
from copy import deepcopy


class LegacyTariffRepository:
    def __init__(self, get_data, save_dataset):
        self.get_data = get_data
        self.save_dataset = save_dataset

    def rows(self):
        return deepcopy(list(self.get_data().get("transportes_tarifarios", []) or []))

    def replace(self, rows, *, expected):
        current = self.get_data()
        if list(current.get("transportes_tarifarios", []) or []) != expected:
            raise ValueError("Os tarifarios foram alterados. Atualize e tente novamente.")
        present = "transportes_tarifarios" in current
        previous = deepcopy(current.get("transportes_tarifarios"))
        current["transportes_tarifarios"] = deepcopy(rows)
        try:
            self.save_dataset(force=True, blocking=True)
        except Exception:
            current = self.get_data()
            if present:
                current["transportes_tarifarios"] = previous
            else:
                current.pop("transportes_tarifarios", None)
            raise
