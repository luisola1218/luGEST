"""Catalog-specific purchase unit of work over the current persistence adapter."""
from copy import deepcopy
from .legacy_note_repository import LegacyNoteRepository
from ..application.note_commands import PurchaseCatalogs


class LegacyPurchaseRepository(LegacyNoteRepository):
    FIELDS = {"notes": "notas_encomenda", "materials": "materiais", "products": "produtos", "assemblies": "conjuntos"}

    def load(self):
        data = self.get_data()
        return PurchaseCatalogs(**{field: deepcopy(list(data.get(key, []) or [])) for field, key in self.FIELDS.items()})

    def save_catalogs(self, catalogs, *, expected):
        self._replace_catalogs({key: getattr(catalogs, field) for field, key in self.FIELDS.items()},
                               {key: getattr(expected, field) for field, key in self.FIELDS.items()})
