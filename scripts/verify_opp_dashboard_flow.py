from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _card_value(cards: list[dict], title: str) -> float:
    for card in cards:
        if str(card.get("title", "") or "").strip() == title:
            raw = str(card.get("value", "") or "").strip().replace(" EUR", "").replace(".", "").replace(",", ".")
            try:
                return float(raw or 0)
            except Exception:
                return 0.0
    return 0.0


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()

    material_id = "MAT00003"
    product_code = "PRD-0002"

    material_before = backend.material_by_id(material_id) or {}
    product_before = next(
        (
            row
            for row in list(data.get("produtos", []) or [])
            if str(row.get("codigo", "") or "").strip() == product_code
        ),
        {},
    )
    material_qty_before = float(material_before.get("quantidade", 0) or 0)
    product_qty_before = float(product_before.get("qty", 0) or 0)

    finance_before = backend.finance_dashboard("2026")
    compras_mp_before = _card_value(finance_before.get("cards", []), "Compras MP")
    compras_prod_before = _card_value(finance_before.get("cards", []), "Compras Produtos")

    created_order = ""
    parent_number = ""
    child_numbers: list[str] = []
    try:
        enc = backend.order_create_or_update(
            {
                "cliente": "CL0002",
                "nota_cliente": "VERIFY_OPP_DASHBOARD",
                "data_entrega": "2026-03-31",
                "tempo_estimado": 30,
            }
        )
        created_order = str(enc.get("numero", "") or "").strip()
        backend.order_piece_create_or_update(
            created_order,
            {
                "material": "AISI 304L 2B",
                "espessura": "2",
                "descricao": "VERIFY OPP DASHBOARD",
                "ref_externa": "VERIFY-OPP-DASH-001",
                "quantidade_pedida": 3,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        opp_rows = backend.opp_rows(created_order, "Todas", "2026", "Todas", "CL0002")
        if len(opp_rows) != 1:
            raise RuntimeError(f"OPP não encontrada de forma unívoca: {opp_rows}")
        opp_row = opp_rows[0]
        if str(opp_row.get("ano", "") or "").strip() != "2026":
            raise RuntimeError(f"Ano OPP incorreto: {opp_row}")
        opp = str(opp_row.get("opp", "") or "").strip()
        if not opp:
            raise RuntimeError(f"OPP sem número: {opp_row}")
        detail = backend.opp_detail(opp)
        if str(detail.get("encomenda", "") or "").strip() != created_order:
            raise RuntimeError(f"Detalhe OPP não corresponde à encomenda: {detail}")
        if len(detail.get("operacoes", []) or []) != 2:
            raise RuntimeError(f"Operações OPP inesperadas: {detail}")

        parent = backend.ne_create_draft()
        parent_number = str(parent.get("numero", "") or "").strip()
        backend.ne_save(
            {
                "numero": parent_number,
                "fornecedor_id": "",
                "fornecedor": "",
                "contacto": "",
                "data_entrega": "2026-03-31",
                "obs": "VERIFY_OPP_DASHBOARD",
                "local_descarga": "Nossas Instalações",
                "meio_transporte": "Nosso Cargo",
                "lines": [
                    {
                        "ref": material_id,
                        "descricao": "AISI 304L 2mm | 3000x1500",
                        "fornecedor_linha": "FOR-0001 - AçoNorte, Lda",
                        "origem": "Matéria-Prima",
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
                        "ref": product_code,
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
                "guia": "GUIA-OPP-DASH",
                "fatura": "FT-OPP-DASH",
                "obs": "VERIFY_OPP_DASHBOARD",
                "lines": [
                    {"ref": material_id, "qtd": 1},
                    {"ref": product_code, "qtd": 2},
                ],
            },
        )

        finance_after = backend.finance_dashboard("2026")
        compras_mp_after = _card_value(finance_after.get("cards", []), "Compras MP")
        compras_prod_after = _card_value(finance_after.get("cards", []), "Compras Produtos")
        if round(compras_mp_after - compras_mp_before, 2) != 140.0:
            raise RuntimeError(f"Compras MP incorretas: antes={compras_mp_before} depois={compras_mp_after}")
        if round(compras_prod_after - compras_prod_before, 2) != 3.0:
            raise RuntimeError(f"Compras Produtos incorretas: antes={compras_prod_before} depois={compras_prod_after}")

        finance_empty = backend.finance_dashboard("2099")
        if _card_value(finance_empty.get("cards", []), "Compras MP") != 0.0:
            raise RuntimeError(f"Filtro anual MP incorreto: {finance_empty.get('cards')}")
        if _card_value(finance_empty.get("cards", []), "Compras Produtos") != 0.0:
            raise RuntimeError(f"Filtro anual produtos incorreto: {finance_empty.get('cards')}")

        if not any(str(row.get("ne", "") or "").strip() == child_number for row in finance_after.get("compras_materias", []) or []):
            raise RuntimeError("NE de matéria-prima não apareceu no dashboard.")
        if not any(str(row.get("ne", "") or "").strip() == child_number for row in finance_after.get("compras_produtos", []) or []):
            raise RuntimeError("NE de produto não apareceu no dashboard.")

        print("opp-dashboard-flow-ok", created_order, opp, child_number)
        return 0
    finally:
        for number in child_numbers:
            try:
                backend.ne_remove(number)
            except Exception:
                pass
        if parent_number:
            try:
                backend.ne_remove(parent_number)
            except Exception:
                pass
        if created_order:
            try:
                backend.order_remove(created_order)
            except Exception:
                pass
        data = backend.ensure_data()
        material_row = next(
            (
                row
                for row in list(data.get("materiais", []) or [])
                if str(row.get("id", "") or "").strip() == material_id
            ),
            None,
        )
        if isinstance(material_row, dict):
            material_row["quantidade"] = material_qty_before
        product_row = next(
            (
                row
                for row in list(data.get("produtos", []) or [])
                if str(row.get("codigo", "") or "").strip() == product_code
            ),
            None,
        )
        if isinstance(product_row, dict):
            product_row["qty"] = product_qty_before
        backend._save(force=True)


if __name__ == "__main__":
    raise SystemExit(main())
