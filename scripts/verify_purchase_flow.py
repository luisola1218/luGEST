from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def main() -> int:
    backend = LegacyBackend()

    material_before = backend.material_by_id("MAT00003")
    product_before = next(
        (
            row
            for row in list(backend.ensure_data().get("produtos", []) or [])
            if str(row.get("codigo", "") or "").strip() == "PRD-0002"
        ),
        None,
    )
    material_qty_before = float((material_before or {}).get("quantidade", 0) or 0)
    product_qty_before = float((product_before or {}).get("qty", 0) or 0)

    parent = backend.ne_create_draft()
    parent_number = str(parent.get("numero", "") or "").strip()
    child_numbers: list[str] = []
    try:
        backend.ne_save(
            {
                "numero": parent_number,
                "fornecedor_id": "",
                "fornecedor": "",
                "contacto": "",
                "data_entrega": "2026-03-31",
                "obs": "VERIFY_FLOW",
                "local_descarga": "Nossas Instalações",
                "meio_transporte": "Nosso Cargo",
                "lines": [
                    {
                        "ref": "MAT00003",
                        "descricao": "AISI 304L 2mm | 3000x1500",
                        "fornecedor_linha": "FOR-0001 - AçoNorte, Lda",
                        "origem": "Matéria-Prima",
                        "qtd": 1,
                        "unid": "UN",
                        "preco": 140.0,
                        "desconto": 0,
                        "iva": 23,
                    },
                    {
                        "ref": "PRD-0002",
                        "descricao": "Óculos Proteo Incolor",
                        "fornecedor_linha": "FOR-0001 - AçoNorte, Lda",
                        "origem": "Produto",
                        "qtd": 2,
                        "unid": "UN",
                        "preco": 1.5,
                        "desconto": 0,
                        "iva": 23,
                    },
                ],
            }
        )
        backend.ne_approve(parent_number)
        created = backend.ne_generate_supplier_orders(parent_number)
        child_numbers = [str(row.get("numero", "") or "").strip() for row in created]
        if len(child_numbers) != 1:
            raise RuntimeError(f"Geração inesperada de NEs: {created}")
        child_number = child_numbers[0]
        backend.ne_register_delivery(
            child_number,
            {
                "data_entrega": "2026-03-31",
                "data_documento": "2026-03-31",
                "guia": "GUIA-VERIFY",
                "fatura": "FT-VERIFY",
                "obs": "VERIFY_FLOW",
                "lines": [
                    {"ref": "MAT00003", "qtd": 1},
                    {"ref": "PRD-0002", "qtd": 2},
                ],
            },
        )

        child = backend.ne_detail(child_number)
        if not all("ENTREGUE" in str(line.get("entrega", "") or "").upper() for line in child.get("lines", [])):
            raise RuntimeError(f"Entrega não ficou concluída: {child}")

        material_after = backend.material_by_id("MAT00003")
        product_after = next(
            (
                row
                for row in list(backend.ensure_data().get("produtos", []) or [])
                if str(row.get("codigo", "") or "").strip() == "PRD-0002"
            ),
            None,
        )
        material_delta = float((material_after or {}).get("quantidade", 0) or 0) - material_qty_before
        product_delta = float((product_after or {}).get("qty", 0) or 0) - product_qty_before
        if abs(material_delta - 1.0) > 1e-6:
            raise RuntimeError(f"Stock matéria-prima incorreto: delta={material_delta}")
        if abs(product_delta - 2.0) > 1e-6:
            raise RuntimeError(f"Stock produto incorreto: delta={product_delta}")

        print("purchase-flow-ok", parent_number, child_number, material_delta, product_delta)
        return 0
    finally:
        for number in child_numbers:
            try:
                backend.ne_remove(number)
            except Exception:
                pass
        try:
            backend.ne_remove(parent_number)
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
