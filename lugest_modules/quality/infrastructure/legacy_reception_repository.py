"""Persist prepared reception catalogs and all associated events once."""
from copy import deepcopy
from lugest_modules.quality.application.receptions import ReceptionCatalogs

class LegacyReceptionRepository:
    fields = {"materials": "materiais", "products": "produtos", "notes": "notas_encomenda",
              "nonconformities": "quality_nonconformities", "documents": "quality_documents"}
    logs = ("stock_log", "produtos_mov", "audit_log")

    def __init__(self, get_data, save, audit, log_stock, product_movement):
        self.get_data, self.save, self.audit = get_data, save, audit
        self.log_stock, self.product_movement = log_stock, product_movement

    def load(self):
        data = self.get_data()
        return ReceptionCatalogs(**{name: deepcopy(data.get(key, []) or []) for name, key in self.fields.items()})

    def replace(self, catalogs, *, expected, audit_events, stock_events, product_events):
        if self.load() != expected:
            raise ValueError("Os registos foram alterados. Atualize e tente novamente.")
        data = self.get_data()
        keys = (*self.fields.values(), *self.logs)
        previous = {key: deepcopy(data[key]) for key in keys if key in data}
        prepared = {key: deepcopy(getattr(catalogs, name)) for name, key in self.fields.items()}
        for key in self.logs:
            if key in data:
                prepared[key] = deepcopy(data[key])
        for action, details, actor in stock_events:
            self.log_stock(prepared, action, details, operador=actor)
        for event in product_events:
            self.product_movement(prepared, **deepcopy(event))
        for event in audit_events:
            self.audit(prepared, **deepcopy(event))
        try:
            data.update(prepared)
            self.save(force=True, audit=False, blocking=True)
        except Exception:
            current = self.get_data()
            for key in keys:
                if key in previous:
                    current[key] = previous[key]
                else:
                    current.pop(key, None)
            raise
