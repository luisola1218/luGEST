"""Publish the quality release catalogs, stock event and audit together."""
from copy import deepcopy
from lugest_modules.quality.application.material_release import ReleaseCatalogs


class LegacyReleaseRepository:
    fields = {"materials": "materiais", "nonconformities": "quality_nonconformities", "notes": "notas_encomenda"}

    def __init__(self, get_data, save, audit, log_stock):
        self.get_data, self.save, self.audit, self.log_stock = get_data, save, audit, log_stock

    def load(self):
        data = self.get_data()
        return ReleaseCatalogs(**{name: deepcopy(data.get(key, []) or []) for name, key in self.fields.items()})

    def replace(self, catalogs, *, expected, event, stock_event):
        if self.load() != expected:
            raise ValueError("Os registos foram alterados. Atualize e tente novamente.")
        data = self.get_data()
        keys = (*self.fields.values(), "stock_log", "audit_log")
        previous = {key: deepcopy(data[key]) for key in keys if key in data}
        prepared = {key: deepcopy(getattr(catalogs, name)) for name, key in self.fields.items()}
        for key in ("stock_log", "audit_log"):
            if key in data:
                prepared[key] = deepcopy(data[key])
        if stock_event:
            action, details, actor = stock_event
            self.log_stock(prepared, action, details, operador=actor)
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
