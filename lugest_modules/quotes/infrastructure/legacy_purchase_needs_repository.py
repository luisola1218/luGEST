"""Read-only stock and quote projections for purchase planning."""
from copy import deepcopy


class LegacyPurchaseNeedsRepository:
    def __init__(self, get_data, material_by_id):
        self.get_data = get_data
        self.material_by_id = material_by_id

    def quote_lines(self, number):
        quote = next((row for row in self.get_data().get("orcamentos", [])
                      if str(row.get("numero", "") or "").strip() == number), None)
        if quote is None:
            raise ValueError("Orcamento nao encontrado.")
        return deepcopy(list(quote.get("linhas", []) or []))

    def products(self):
        fields = ("codigo", "descricao", "qty", "unid", "p_compra")
        return [{key: deepcopy(row[key]) for key in fields if key in row}
                for row in list(self.get_data().get("produtos", []) or [])]

    def material(self, code):
        return deepcopy(self.material_by_id(code))
