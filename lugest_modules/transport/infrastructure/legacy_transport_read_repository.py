"""Read-only transport projections over the historical snapshot."""
from copy import deepcopy


class LegacyTransportReadRepository:
    def __init__(self, get_data):
        self.get_data = get_data

    def trips(self):
        return deepcopy(list(self.get_data().get("transportes", []) or []))

    def orders(self):
        return deepcopy(list(self.get_data().get("encomendas", []) or []))

    def clients(self):
        return deepcopy(list(self.get_data().get("clientes", []) or []))

    def quotes(self):
        fields = ("numero", "zona_transporte", "nota_transporte")
        return [{key: deepcopy(row[key]) for key in fields if key in row}
                for row in self.get_data().get("orcamentos", []) or []]

    def tariffs(self):
        return [{"zona": row.get("zona", "")} for row in self.get_data().get("transportes_tarifarios", []) or []]

    def guides(self):
        return deepcopy(list(self.get_data().get("expedicoes", []) or []))

    def suppliers(self):
        rows = [{"id": str(row.get("id", "") or "").strip(), "nome": str(row.get("nome", "") or "").strip()}
                for row in self.get_data().get("fornecedores", []) or []]
        def key(row):
            code = row["id"].upper()
            return (int(code[4:]) if code.startswith("FOR-") and code[4:].isdigit() else 10**9, code)
        return sorted(rows, key=key)

    def _find(self, collection, key, value):
        return deepcopy(next((row for row in self.get_data().get(collection, []) or []
                              if str(row.get(key, "") or "").strip() == str(value or "").strip()), None))

    def trip(self, number):
        return self._find("transportes", "numero", number)

    def order(self, number):
        return self._find("encomendas", "numero", number)

    def quote(self, number):
        return next((row for row in self.quotes() if str(row.get("numero", "") or "").strip() == str(number or "").strip()), None)

    def client(self, code):
        return self._find("clientes", "codigo", code) or {}
