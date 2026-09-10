"""Create a supplier quotation from calculated shortages through a purchasing port."""
from typing import Any, Callable
from .purchase_needs import PurchaseNeeds


class PurchaseQuote:
    def __init__(self, needs: PurchaseNeeds, parse_float: Callable, save_note: Callable, note_detail: Callable):
        self.needs = needs
        self.parse_float = parse_float
        self.save_note = save_note
        self.note_detail = note_detail

    def create(self, numero: str, lines: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        needs = self.needs.rows(numero_txt, lines)
        if not needs:
            raise ValueError("Nao existem necessidades de compra nas linhas do orcamento.")
        note = self.save_note(
            {
                "fornecedor": "",
                "fornecedor_id": "",
                "contacto": "",
                "obs": f"Pedido de cotacao gerado a partir do orcamento {numero_txt}".strip(),
                "lines": [
                    (
                        {
                            "ref": str(need.get("ref", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip() or str(need.get("material", "") or "").strip(),
                            "origem": "Materia-prima",
                            "qtd": round(self.parse_float(need.get("qtd", 0), 0), 2),
                            "unid": str(need.get("unid", "") or "UN").strip() or "UN",
                            "preco": round(self.parse_float(need.get("preco", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "material": str(need.get("material", "") or "").strip(),
                            "espessura": str(need.get("espessura", "") or "").strip(),
                            "dimensao": str(need.get("dimensao", "") or "").strip(),
                            "dimensoes": str(need.get("dimensao", "") or "").strip(),
                            "formato": str(need.get("formato", "") or "Chapa").strip() or "Chapa",
                            "comprimento": self.parse_float(need.get("comprimento", 0), 0),
                            "largura": self.parse_float(need.get("largura", 0), 0),
                            "diametro": self.parse_float(need.get("diametro", 0), 0),
                            "metros": self.parse_float(need.get("metros", 0), 0),
                            "kg_m": self.parse_float(need.get("kg_m", 0), 0),
                            "peso_unid": self.parse_float(need.get("peso_unid", 0), 0),
                            "_material_pending_create": bool(need.get("_material_pending_create", False)),
                            "_material_manual": bool(need.get("_material_manual", False)),
                        }
                        if str(need.get("kind", "") or "") == "material"
                        else {
                            "ref": str(need.get("ref", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip(),
                            "origem": "Produto",
                            "qtd": round(self.parse_float(need.get("qtd", 0), 0), 2),
                            "unid": str(need.get("unid", "") or "UN").strip() or "UN",
                            "preco": round(self.parse_float(need.get("preco", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "_product_pending_create": bool(need.get("_product_pending_create", False)),
                        }
                    )
                    for need in needs
                ],
            }
        )
        note_number = str(note.get("numero", "") or "").strip()
        return {"numero": note_number, "line_count": len(list(note.get("linhas", []) or [])), "needs": needs, "detail": self.note_detail(note_number)}
