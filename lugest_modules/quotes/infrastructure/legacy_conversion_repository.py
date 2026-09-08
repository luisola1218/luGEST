"""Stage conversion changes away from the shared runtime snapshot.

Identifier allocators may reserve SQL counters independently. This adapter
protects local aggregate preparation and immediate save failures; it does not
claim one database transaction across all historical integrations.
"""
from copy import deepcopy


class LegacyConversionRepository:
    FIELDS = ("clientes", "encomendas", "orcamentos", "seq", "refs", "of_seq", "opp_seq")

    def __init__(self, get_data, save_dataset, *, next_client, next_order,
                 order_code, make_line_ports, register_reference):
        self.get_data = get_data
        self.save_dataset = save_dataset
        self.next_client = next_client
        self.next_order = next_order
        self.order_code = order_code
        self.make_line_ports = make_line_ports
        self.register_reference = register_reference
        self.candidate = deepcopy(get_data())
        self.original = {key: deepcopy(self.candidate[key]) for key in self.FIELDS if key in self.candidate}

    def quote(self, number):
        return deepcopy(next((row for row in self.candidate.get("orcamentos", [])
                              if str(row.get("numero", "") or "").strip() == number), None))

    def clients(self):
        return deepcopy(list(self.candidate.get("clientes", []) or []))

    def allocate_client_code(self):
        return self.next_client(self.candidate)

    def add_client(self, client):
        self.candidate.setdefault("clientes", []).append(deepcopy(client))

    def allocate_order_number(self):
        return self.next_order(self.candidate)

    def allocate_order_code(self, order):
        return self.order_code(self.candidate, order)

    def line_ports(self, client_code):
        return self.make_line_ports(self.candidate, client_code)

    def add_reference(self, internal, external):
        self.register_reference(self.candidate, internal, external)

    def save(self, quote, order):
        current = self.get_data()
        for key in self.FIELDS:
            if (key in current) != (key in self.original) or current.get(key) != self.original.get(key):
                raise ValueError("Os dados da conversao foram alterados. Atualize e tente novamente.")
        self.candidate.setdefault("encomendas", []).append(deepcopy(order))
        number = str(quote.get("numero", "") or "").strip()
        target = next(row for row in self.candidate.get("orcamentos", [])
                      if str(row.get("numero", "") or "").strip() == number)
        target.update(deepcopy(quote))
        for key in self.FIELDS:
            if key in self.candidate:
                current[key] = deepcopy(self.candidate[key])
        try:
            self.save_dataset(force=True, blocking=True)
        except Exception:
            current = self.get_data()
            for key in self.FIELDS:
                if key in self.original:
                    current[key] = deepcopy(self.original[key])
                else:
                    current.pop(key, None)
            raise
