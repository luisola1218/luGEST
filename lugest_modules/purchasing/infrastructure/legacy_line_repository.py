"""Detached catalog reads required by purchase line validation."""
from copy import deepcopy


class LegacyLineRepository:
    def __init__(self, get_data):
        self.get_data = get_data

    def _find(self, collection, key, code):
        return deepcopy(next((row for row in self.get_data().get(collection, []) or []
                              if str(row.get(key, "") or "").strip() == str(code or "").strip()), None))

    def material(self, code):
        return self._find("materiais", "id", code)

    def product(self, code):
        return self._find("produtos", "codigo", code)
