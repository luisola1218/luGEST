"""Trip writes and derived order links, with immediate failure recovery."""
from copy import deepcopy
from lugest_modules.transport.application.order_links import synchronize_orders


class LegacyTripRepository:
    def __init__(self, get_data, save_dataset, norm_text):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.norm_text = norm_text

    def get(self, number):
        return deepcopy(next((row for row in self.get_data().get("transportes", [])
                              if str(row.get("numero", "") or "").strip() == number), None))

    def save(self, trip, *, expected, sync_links=False):
        data = self.get_data()
        number = str(trip.get("numero", "") or "").strip()
        rows = deepcopy(list(data.get("transportes", []) or []))
        index = next((i for i, row in enumerate(rows) if str(row.get("numero", "") or "").strip() == number), None)
        if index is None or rows[index] != expected:
            raise ValueError("A viagem foi alterada. Atualize e tente novamente.")
        rows[index] = deepcopy(trip)
        fields = ("transportes", "encomendas") if sync_links else ("transportes",)
        previous = {key: deepcopy(data[key]) for key in fields if key in data}
        orders = synchronize_orders(rows, list(data.get("encomendas", []) or []), self.norm_text) if sync_links else []
        data["transportes"] = rows
        if sync_links:
            data["encomendas"] = orders
        try:
            self.save_dataset(force=True, blocking=True)
        except Exception:
            current = self.get_data()
            for key in fields:
                if key in previous:
                    current[key] = previous[key]
                else:
                    current.pop(key, None)
            raise
