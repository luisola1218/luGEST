from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def main() -> None:
    backend = LegacyBackend()
    data = backend.ensure_data()
    supplier_id = "FOR-TEST-GUARD"
    note_number = "NE-TEST-FK-GUARD"
    supplier = {
        "id": supplier_id,
        "nome": "Fornecedor Guard",
        "contacto": "999999999",
        "email": "",
        "nif": "",
        "morada": "",
        "codigo_postal": "",
        "localidade": "",
        "pais": "",
        "cond_pagamento": "",
        "prazo_entrega_dias": 0,
        "website": "",
        "obs": "",
    }
    before_suppliers = list(data.get("fornecedores", []) or [])
    before_notes = list(data.get("notas_encomenda", []) or [])
    try:
        data.setdefault("fornecedores", []).append(supplier)
        saved = backend.ne_save(
            {
                "numero": note_number,
                "fornecedor_id": "FOR-0000",
                "fornecedor": f"{supplier_id} - {supplier['nome']}",
                "contacto": "",
                "data_entrega": "",
                "obs": "teste fk fornecedor",
                "local_descarga": "",
                "meio_transporte": "",
                "lines": [
                    {
                        "ref": "TEST-FK-L1",
                        "descricao": "Linha teste FK",
                        "fornecedor_linha": f"{supplier_id} - {supplier['nome']}",
                        "origem": "Produto",
                        "qtd": 1,
                        "unid": "UN",
                        "preco": 1,
                        "desconto": 0,
                        "iva": 23,
                    }
                ],
            }
        )
        assert str(saved.get("fornecedor_id", "") or "").strip() == supplier_id, saved
        detail = backend.ne_detail(note_number)
        assert str(detail.get("fornecedor_id", "") or "").strip() == supplier_id, detail
        print("ne-supplier-fk-guard-ok")
    finally:
        data["notas_encomenda"] = before_notes
        data["fornecedores"] = before_suppliers
        backend._save(force=True)


if __name__ == "__main__":
    main()
