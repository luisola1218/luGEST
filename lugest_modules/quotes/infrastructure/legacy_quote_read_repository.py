"""Read detached quote aggregates; never expose the global dataset."""
from copy import deepcopy
from typing import Callable


class LegacyQuoteReadRepository:
    SUMMARY_FIELDS = ('numero', 'cliente', 'estado', 'numero_encomenda', 'total', 'data', 'ano')

    def __init__(self, quotes: Callable[[], list[dict]]):
        self.quotes = quotes

    def summaries(self) -> list[dict]:
        rows = []
        for quote in self.quotes() or []:
            if not isinstance(quote, dict):
                continue
            row = {key: deepcopy(quote[key]) for key in self.SUMMARY_FIELDS if key in quote}
            row['line_count'] = len(quote.get('linhas', []) or [])
            rows.append(row)
        return rows

    def get(self, number: str) -> dict | None:
        return deepcopy(next((row for row in self.quotes()
                              if str(row.get('numero', '') or '').strip() == number), None))
