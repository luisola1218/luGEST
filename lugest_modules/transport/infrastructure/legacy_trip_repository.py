"""Trip writes and derived order links, with immediate failure recovery."""
from copy import deepcopy
from lugest_modules.transport.application.order_links import synchronize_orders


class LegacyTripRepository:
    def __init__(self, get_data, save_dataset, norm_text, allocate=None):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.norm_text = norm_text
        self.allocate = allocate
        self.pending_sequence = None

    def allocate_number(self, requested):
        candidate = deepcopy(self.get_data())
        original = deepcopy(candidate.get("seq"))
        present = "seq" in candidate
        number = self.allocate(candidate, requested)
        self.pending_sequence = (present, original, deepcopy(candidate.get("seq")))
        return number

    def get(self, number):
        return deepcopy(next((row for row in self.get_data().get("transportes", [])
                              if str(row.get("numero", "") or "").strip() == number), None))

    def all(self):
        return deepcopy(list(self.get_data().get("transportes", []) or []))

    def order(self, number):
        return deepcopy(next((row for row in self.get_data().get("encomendas", [])
                              if str(row.get("numero", "") or "").strip() == number), None))

    def client(self, code):
        return deepcopy(next((row for row in self.get_data().get("clientes", [])
                              if str(row.get("codigo", "") or "").strip() == code), {}))

    def save_assignment(self, trip, *, expected, catalog):
        if self.all() != catalog:
            raise ValueError("As viagens foram alteradas. Atualize e tente novamente.")
        return self.save(trip, expected=expected, sync_links=True)

    def save(self, trip, *, expected, sync_links=False):
        return self._write(trip, expected=expected, sync_links=sync_links)

    def remove(self, number, *, expected):
        if str(expected.get("numero", "") or "").strip() != number:
            raise ValueError("Transporte invalido.")
        return self._write(None, expected=expected, sync_links=True)

    def _write(self, trip, *, expected, sync_links):
        data = self.get_data()
        number = str((trip or expected).get("numero", "") or "").strip()
        rows = deepcopy(list(data.get("transportes", []) or []))
        index = next((i for i, row in enumerate(rows) if str(row.get("numero", "") or "").strip() == number), None)
        current = rows[index] if index is not None else None
        if current != expected:
            raise ValueError("A viagem foi alterada. Atualize e tente novamente.")
        if trip is None:
            rows.pop(index)
        elif index is None:
            rows.append(deepcopy(trip))
        else:
            rows[index] = deepcopy(trip)
        fields = ("transportes", "encomendas") if sync_links else ("transportes",)
        sequence = self.pending_sequence
        if sequence is not None:
            present, original, updated = sequence
            if ("seq" in data) != present or data.get("seq") != original:
                raise ValueError("A sequencia foi alterada. Atualize e tente novamente.")
            fields += ("seq",)
        previous = {key: deepcopy(data[key]) for key in fields if key in data}
        orders = synchronize_orders(rows, list(data.get("encomendas", []) or []), self.norm_text) if sync_links else []
        data["transportes"] = rows
        if sequence is not None:
            data["seq"] = updated
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
        finally:
            self.pending_sequence = None
