from __future__ import annotations

import copy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _product_by_code(backend: LegacyBackend, code: str) -> dict | None:
    return next(
        (
            row
            for row in list(backend.ensure_data().get("produtos", []) or [])
            if str(row.get("codigo", "") or "").strip() == code
        ),
        None,
    )


def _restore_record(rows: list, key_name: str, key_value: str, snapshot: dict) -> None:
    if not snapshot:
        return
    for index, row in enumerate(list(rows or [])):
        if str((row or {}).get(key_name, "") or "").strip() == key_value:
            rows[index] = copy.deepcopy(snapshot)
            return


def main() -> int:
    backend = LegacyBackend()

    material_before = copy.deepcopy(backend.material_by_id("MAT00003") or {})
    product_before = copy.deepcopy(_product_by_code(backend, "PRD-0002") or {})
    created_material = not bool(material_before)
    created_product = not bool(product_before)
    if created_material:
        backend.ensure_data().setdefault("materiais", []).append(
            {
                "id": "MAT00003",
                "material": "Inox 304L",
                "espessura": "2",
                "formato": "Chapa",
                "dimensao": "3000x1500",
                "quantidade": 0,
                "reservado": 0,
                "quality_received_qty": 0,
                "quality_status": "",
            }
        )
        material_before = copy.deepcopy(backend.material_by_id("MAT00003") or {})
    if created_product:
        backend.ensure_data().setdefault("produtos", []).append(
            {
                "codigo": "PRD-0002",
                "descricao": "Oculos protecao incolor",
                "unid": "UN",
                "qty": 0,
                "quality_received_qty": 0,
                "quality_status": "",
            }
        )
        product_before = copy.deepcopy(_product_by_code(backend, "PRD-0002") or {})
    if created_material or created_product:
        backend._save(force=True)
    material_qty_before = float(material_before.get("quantidade", 0) or 0)
    product_qty_before = float(product_before.get("qty", 0) or 0)
    material_received_before = float(material_before.get("quality_received_qty", 0) or 0)
    product_received_before = float(product_before.get("quality_received_qty", 0) or 0)

    parent = backend.ne_create_draft()
    parent_number = str(parent.get("numero", "") or "").strip()
    child_numbers: list[str] = []
    try:
        backend.ne_save(
            {
                "numero": parent_number,
                "fornecedor_id": "FOR-0001",
                "fornecedor": "Luis",
                "contacto": "",
                "data_entrega": "2026-03-31",
                "obs": "VERIFY_FLOW",
                "local_descarga": "Nossas Instalacoes",
                "meio_transporte": "Nosso Cargo",
                "lines": [
                    {
                        "ref": "MAT00003",
                        "descricao": "AISI 304L 2mm | 3000x1500",
                        "fornecedor_linha": "FOR-0001 - Luis",
                        "origem": "Materia-Prima",
                        "material": "Inox 304L",
                        "espessura": "2",
                        "formato": "Chapa",
                        "dimensao": "3000x1500",
                        "qtd": 1,
                        "unid": "UN",
                        "preco": 140.0,
                        "desconto": 0,
                        "iva": 23,
                    },
                    {
                        "ref": "PRD-0002",
                        "descricao": "Oculos protecao incolor",
                        "fornecedor_linha": "FOR-0001 - Luis",
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
            raise RuntimeError(f"Geracao inesperada de NEs: {created}")
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
            raise RuntimeError(f"Entrega nao ficou concluida: {child}")

        material_after = backend.material_by_id("MAT00003")
        product_after = _product_by_code(backend, "PRD-0002")
        material_delta = float((material_after or {}).get("quantidade", 0) or 0) - material_qty_before
        product_delta = float((product_after or {}).get("qty", 0) or 0) - product_qty_before
        material_received_delta = float((material_after or {}).get("quality_received_qty", 0) or 0) - material_received_before
        product_received_delta = float((product_after or {}).get("quality_received_qty", 0) or 0) - product_received_before
        if abs(material_delta) > 1e-6:
            raise RuntimeError(f"Stock disponivel de materia-prima mudou antes da qualidade: delta={material_delta}")
        if abs(product_delta) > 1e-6:
            raise RuntimeError(f"Stock disponivel de produto mudou antes da qualidade: delta={product_delta}")
        if abs(material_received_delta - 1.0) > 1e-6:
            raise RuntimeError(f"Rececao qualidade materia-prima incorreta: delta={material_received_delta}")
        if abs(product_received_delta - 2.0) > 1e-6:
            raise RuntimeError(f"Rececao qualidade produto incorreta: delta={product_received_delta}")
        if str((material_after or {}).get("quality_status", "") or "").strip() != "EM_INSPECAO":
            raise RuntimeError(f"Materia-prima nao ficou em inspecao: {material_after}")
        if str((product_after or {}).get("quality_status", "") or "").strip() != "EM_INSPECAO":
            raise RuntimeError(f"Produto nao ficou em inspecao: {product_after}")

        print("purchase-flow-ok", parent_number, child_number, material_received_delta, product_received_delta)
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
        data = backend.ensure_data()
        if created_material:
            data["materiais"] = [row for row in list(data.get("materiais", []) or []) if str((row or {}).get("id", "") or "") != "MAT00003"]
        else:
            _restore_record(data.setdefault("materiais", []), "id", "MAT00003", material_before)
        if created_product:
            data["produtos"] = [row for row in list(data.get("produtos", []) or []) if str((row or {}).get("codigo", "") or "") != "PRD-0002"]
        else:
            _restore_record(data.setdefault("produtos", []), "codigo", "PRD-0002", product_before)
        backend._save(force=True)


if __name__ == "__main__":
    raise SystemExit(main())
