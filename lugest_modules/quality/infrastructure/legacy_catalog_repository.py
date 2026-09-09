"""Persist a prepared quality catalog and its audit event with local recovery."""
from copy import deepcopy


class LegacyQualityCatalogRepository:
    def __init__(self, get_data, save_dataset, append_audit, collection):
        if collection not in {"quality_nonconformities", "quality_documents"}:
            raise ValueError("Unknown quality catalog")
        self.collection = collection
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.append_audit = append_audit

    def rows(self):
        return deepcopy(list(self.get_data().get(self.collection, []) or []))

    def replace(self, rows, *, expected, event=None):
        data = self.get_data()
        if list(data.get(self.collection, []) or []) != expected:
            raise ValueError("Os registos de qualidade foram alterados. Atualize e tente novamente.")
        fields = (self.collection, "audit_log")
        previous = {key: deepcopy(data[key]) for key in fields if key in data}
        data[self.collection] = deepcopy(rows)
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
