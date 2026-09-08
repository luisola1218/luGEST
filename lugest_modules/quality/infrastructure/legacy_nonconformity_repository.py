"""Persist prepared nonconformities and their audit event together in the cache."""
from copy import deepcopy


class LegacyNonconformityRepository:
    def __init__(self, get_data, save_dataset, append_audit):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.append_audit = append_audit

    def rows(self):
        return deepcopy(list(self.get_data().get("quality_nonconformities", []) or []))

    def replace(self, rows, *, expected, event=None):
        data = self.get_data()
        if list(data.get("quality_nonconformities", []) or []) != expected:
            raise ValueError("As nao conformidades foram alteradas. Atualize e tente novamente.")
        fields = ("quality_nonconformities", "audit_log")
        previous = {key: deepcopy(data[key]) for key in fields if key in data}
        data["quality_nonconformities"] = deepcopy(rows)
        try:
            if event is not None:
                self.append_audit(data, **deepcopy(event))
            self.save_dataset(force=True, audit=False, blocking=True)
        except Exception:
            current = self.get_data()
            for key in fields:
                if key in previous:
                    current[key] = previous[key]
                else:
                    current.pop(key, None)
            raise
