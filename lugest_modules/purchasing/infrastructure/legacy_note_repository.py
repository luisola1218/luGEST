"""Persist prepared purchase notes and local numbering with immediate recovery."""
from copy import deepcopy


class LegacyNoteRepository:
    def __init__(self, get_data, save_dataset, allocate):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.allocate = allocate
        self.candidate = None
        self.original_sequence = None

    def rows(self):
        return deepcopy(list(self.get_data().get("notas_encomenda", []) or []))

    def allocate_number(self):
        if self.candidate is None:
            data = self.get_data()
            self.original_sequence = ("seq" in data, deepcopy(data.get("seq")))
            self.candidate = {"notas_encomenda": self.rows(), "seq": deepcopy(data.get("seq", {}))}
        return self.allocate(self.candidate)

    def replace(self, rows, *, expected):
        return self._replace_catalogs({"notas_encomenda": rows}, {"notas_encomenda": expected})

    def _replace_catalogs(self, updates, expected):
        data = self.get_data()
        try:
            if any(list(data.get(key, []) or []) != rows for key, rows in expected.items()):
                raise ValueError("As notas foram alteradas. Atualize e tente novamente.")
            keys = list(updates)
            if self.candidate is not None:
                if self.original_sequence != ("seq" in data, data.get("seq")):
                    raise ValueError("A numeracao foi alterada. Atualize e tente novamente.")
                keys.append("seq")
            previous = {key: deepcopy(data[key]) for key in keys if key in data}
            prepared = deepcopy(updates)
            data.update(prepared)
            if self.candidate is not None:
                data["seq"] = deepcopy(self.candidate["seq"])
            try:
                self.save_dataset(force=True, blocking=True)
            except Exception:
                current = self.get_data()
                for key in keys:
                    if key in previous:
                        current[key] = previous[key]
                    else:
                        current.pop(key, None)
                raise
        finally:
            self.candidate = None
            self.original_sequence = None
