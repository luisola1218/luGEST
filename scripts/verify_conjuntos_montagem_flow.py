from __future__ import annotations

import copy
import sys
import tempfile
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _assert(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def _pick_client(backend: LegacyBackend, suffix: str) -> tuple[dict[str, str], str]:
    data = backend.ensure_data()
    for row in list(data.get("clientes", []) or []):
        if not isinstance(row, dict):
            continue
        code = str(row.get("codigo", "") or "").strip()
        if not code:
            continue
        return (
            {
                "codigo": code,
                "nome": str(row.get("nome", "") or "").strip(),
                "nif": str(row.get("nif", "") or "").strip(),
                "morada": str(row.get("morada", "") or "").strip(),
                "contacto": str(row.get("contacto", "") or "").strip(),
                "email": str(row.get("email", "") or "").strip(),
            },
            "",
        )
    temp_code = f"CLTST{suffix[:4]}"
    payload = {
        "codigo": temp_code,
        "nome": f"Cliente Verify {suffix}",
        "nif": "",
        "morada": "",
        "contacto": "",
        "email": "",
    }
    data.setdefault("clientes", []).append(dict(payload))
    return payload, temp_code


def _pick_material(backend: LegacyBackend) -> tuple[str, str]:
    for row in list(backend.ensure_data().get("materiais", []) or []):
        if not isinstance(row, dict):
            continue
        material = str(row.get("material", "") or "").strip()
        espessura = str(row.get("espessura", "") or "").strip()
        if material and espessura:
            return material, espessura
    return "AISI 304L", "2"


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    token = uuid.uuid4().hex[:8].upper()
    product_code = f"TST-CJ-{token}"
    model_code = f"CJ-TST-{token}"
    quote_num = f"ORC-VERIFY-CJ-{token}"
    assembly_only_quote_num = f"ORC-VERIFY-MONT-{token}"
    temp_ref_ext = f"EXT-CJ-{token}"
    pdf_path = Path(tempfile.gettempdir()) / f"lugest_verify_conjuntos_{token}.pdf"
    supplier_id = backend.supplier_next_id()
    supplier_name = f"Fornecedor Verify {token}"
    supplier_text = f"{supplier_id} - {supplier_name}"
    acceptable_supplier_values = {supplier_text.strip(), supplier_name.strip()}
    purchase_history_num = ""
    purchase_note_num = ""

    client_payload, created_client_code = _pick_client(backend, token)
    client_code = str(client_payload.get("codigo", "") or "").strip()
    material_name, espessura = _pick_material(backend)

    seq_snapshot = copy.deepcopy(dict(data.get("seq", {}) or {}))
    of_seq_snapshot = data.get("of_seq", 1)
    opp_seq_snapshot = data.get("opp_seq", 1)
    orc_seq_snapshot = data.get("orc_seq", 1)

    order_num = ""
    assembly_only_order_num = ""
    cleanup_errors: list[str] = []

    try:
        backend.product_save(
            {
                "codigo": product_code,
                "descricao": f"Parafuso montagem {token}",
                "categoria": "Fixacao",
                "tipo": "Consumivel",
                "unid": "UN",
                "qty": 20,
                "alerta": 2,
                "p_compra": 0.12,
                "obs": "VERIFY_CONJUNTOS_MONTAGEM",
            }
        )
        backend.supplier_save(
            {
                "id": supplier_id,
                "nome": supplier_name,
                "contacto": "verify@supplier.test",
            }
        )
        purchase_history = backend.ne_save(
            {
                "fornecedor": supplier_text,
                "fornecedor_id": supplier_id,
                "contacto": "verify@supplier.test",
                "data_entrega": "2026-03-20",
                "obs": "HISTORICO VERIFY MONTAGEM",
                "lines": [
                    {
                        "ref": product_code,
                        "descricao": f"Parafuso montagem {token}",
                        "fornecedor_linha": supplier_text,
                        "origem": "Produto",
                        "qtd": 5,
                        "unid": "UN",
                        "preco": 0.12,
                        "desconto": 0.0,
                        "iva": 23.0,
                    }
                ],
            }
        )
        purchase_history_num = str(purchase_history.get("numero", "") or "").strip()

        backend.assembly_model_save(
            {
                "codigo": model_code,
                "descricao": f"Conjunto teste {token}",
                "notas": "VERIFY_CONJUNTOS_MONTAGEM",
                "ativo": True,
                "itens": [
                    {
                        "tipo_item": "peca_fabricada",
                        "ref_externa": temp_ref_ext,
                        "descricao": "Painel lateral",
                        "material": material_name,
                        "espessura": espessura,
                        "operacao": "Corte Laser + Quinagem",
                        "qtd": 2,
                        "tempo_peca_min": 7.5,
                        "preco_unit": 35.0,
                    },
                    {
                        "tipo_item": "produto_stock",
                        "produto_codigo": product_code,
                        "descricao": "Parafuso M6",
                        "qtd": 4,
                        "preco_unit": 0.12,
                    },
                    {
                        "tipo_item": "servico_montagem",
                        "descricao": "Montagem final do conjunto",
                        "produto_unid": "SV",
                        "qtd": 1,
                        "tempo_peca_min": 18,
                        "preco_unit": 25.0,
                    },
                ],
            }
        )

        expanded = backend.assembly_model_expand(model_code, quantity=2)
        _assert(len(expanded) == 3, f"Expansao inesperada do conjunto: {expanded}")
        _assert(all(str(row.get("conjunto_codigo", "") or "").strip() == model_code for row in expanded), "Conjunto nao foi propagado nas linhas.")
        _assert(len({str(row.get("grupo_uuid", "") or "").strip() for row in expanded}) == 1, "Grupo de expansao nao ficou consistente.")
        product_line = next(row for row in expanded if str(row.get("tipo_item", "")) == "produto_stock")
        service_line = next(row for row in expanded if str(row.get("tipo_item", "")) == "servico_montagem")
        piece_line = next(row for row in expanded if str(row.get("tipo_item", "")) == "peca_fabricada")
        _assert(abs(float(product_line.get("qtd", 0) or 0) - 8.0) < 1e-6, f"Quantidade expandida do produto invalida: {product_line}")
        _assert(abs(float(service_line.get("qtd", 0) or 0) - 2.0) < 1e-6, f"Quantidade expandida do servico invalida: {service_line}")
        _assert(abs(float(piece_line.get("qtd_base", 0) or 0) - 2.0) < 1e-6, f"Quantidade base da peca invalida: {piece_line}")

        quote = backend.orc_save(
            {
                "numero": quote_num,
                "estado": "Aprovado",
                "cliente": client_payload,
                "linhas": expanded,
                "iva_perc": 23,
                "nota_cliente": "VERIFY_CONJUNTOS_MONTAGEM",
                "executado_por": "VERIFY",
                "nota_transporte": "Entrega em obra",
            }
        )
        _assert(str(quote.get("numero", "") or "").strip() == quote_num, f"Orcamento nao guardado corretamente: {quote}")
        _assert(len(list(quote.get("linhas", []) or [])) == 3, "Linhas do orcamento ficaram incompletas.")
        notes = backend.orc_suggest_notes(quote)
        notes_norm = notes.lower()
        _assert("montagem" in notes_norm, f"Notas do orcamento sem montagem: {notes}")
        _assert("stock" in notes_norm, f"Notas do orcamento sem referencia a stock: {notes}")

        rendered = backend.orc_render_pdf(quote_num, pdf_path)
        _assert(rendered.exists(), f"PDF nao foi criado: {rendered}")
        _assert(rendered.stat().st_size > 0, f"PDF vazio: {rendered}")

        converted = backend.orc_convert_to_order(quote_num, nota_cliente="VERIFY_CONJUNTOS_MONTAGEM")
        order = dict(converted.get("encomenda", {}) or {})
        order_num = str(order.get("numero", "") or "").strip()
        _assert(order_num, f"Conversao em encomenda falhou: {converted}")
        _assert(len(list(order.get("pieces", []) or [])) == 1, f"Pecas esperadas na encomenda nao batem certo: {order}")
        montagem_items = list(order.get("montagem_items", []) or [])
        _assert(len(montagem_items) == 2, f"Itens de montagem inesperados: {montagem_items}")
        _assert(any(str(item.get("produto_codigo", "") or "").strip() == product_code for item in montagem_items), "Produto de montagem nao transitou para a encomenda.")
        _assert(any(str(item.get("tipo_item", "") or "").strip() == "servico_montagem" for item in montagem_items), "Servico de montagem nao transitou para a encomenda.")

        product_before = backend.product_detail(product_code)
        stock_before = float(product_before.get("qty", 0) or 0)
        backend.order_consume_montagem(order_num, operador="VERIFY")
        order_after = backend.order_detail(order_num)
        product_after = backend.product_detail(product_code)
        stock_after = float(product_after.get("qty", 0) or 0)
        _assert(abs(stock_after - (stock_before - 8.0)) < 1e-6, f"Baixa de stock de montagem incorreta: antes={stock_before} depois={stock_after}")
        _assert(str(order_after.get("montagem_estado", "") or "").strip() == "Consumida", f"Estado de montagem inesperado: {order_after}")

        order_items_after = list(order_after.get("montagem_items", []) or [])
        product_after_item = next(item for item in order_items_after if str(item.get("produto_codigo", "") or "").strip() == product_code)
        service_after_item = next(item for item in order_items_after if str(item.get("tipo_item", "") or "").strip() == "servico_montagem")
        _assert(abs(float(product_after_item.get("qtd_consumida", 0) or 0) - 8.0) < 1e-6, f"Quantidade consumida do produto invalida: {product_after_item}")
        _assert(str(product_after_item.get("estado", "") or "").strip() == "Consumido", f"Estado do produto de montagem invalido: {product_after_item}")
        _assert(abs(float(service_after_item.get("qtd_consumida", 0) or 0) - 2.0) < 1e-6, f"Quantidade consumida do servico invalida: {service_after_item}")
        _assert(str(service_after_item.get("estado", "") or "").strip() == "Concluido", f"Estado do servico de montagem invalido: {service_after_item}")

        assembly_quote = backend.orc_save(
            {
                "numero": assembly_only_quote_num,
                "estado": "Aprovado",
                "cliente": client_payload,
                "linhas": [
                    {
                        "tipo_item": "produto_stock",
                        "produto_codigo": product_code,
                        "descricao": "Parafuso M6 montagem pura",
                        "produto_unid": "UN",
                        "qtd": 15,
                        "tempo_peca_min": 0,
                        "preco_unit": 0.12,
                    },
                    {
                        "tipo_item": "servico_montagem",
                        "descricao": "Fecho final montagem pura",
                        "produto_unid": "SV",
                        "qtd": 1,
                        "tempo_peca_min": 35,
                        "preco_unit": 30.0,
                    },
                ],
                "iva_perc": 23,
                "nota_cliente": "VERIFY_MONTAGEM_ONLY",
                "executado_por": "VERIFY",
            }
        )
        _assert(str(assembly_quote.get("numero", "") or "").strip() == assembly_only_quote_num, f"Orcamento de montagem pura falhou: {assembly_quote}")
        assembly_converted = backend.orc_convert_to_order(assembly_only_quote_num, nota_cliente="VERIFY_MONTAGEM_ONLY")
        assembly_order = dict(assembly_converted.get("encomenda", {}) or {})
        assembly_only_order_num = str(assembly_order.get("numero", "") or "").strip()
        _assert(assembly_only_order_num, f"Encomenda de montagem pura nao foi criada: {assembly_converted}")

        assembly_detail = backend.order_detail(assembly_only_order_num)
        _assert(str(assembly_detail.get("estado", "") or "").strip() == "Montagem", f"Estado da encomenda de montagem pura devia ser Montagem: {assembly_detail}")
        _assert(float(assembly_detail.get("montagem_tempo_min", 0) or 0) >= 35.0, f"Tempo de montagem nao ficou refletido: {assembly_detail}")
        _assert(not bool(assembly_detail.get("montagem_stock_ready")), f"Montagem devia sinalizar falta de stock: {assembly_detail}")
        _assert(list(assembly_detail.get("montagem_shortages", []) or []), f"Faltas de stock nao foram apanhadas: {assembly_detail}")

        planning_rows = list(backend.planning_pending_rows(operation="Montagem") or [])
        planning_row = next(
            (
                row
                for row in planning_rows
                if str(row.get("numero", "") or "").strip() == assembly_only_order_num
                and str(row.get("material", "") or "").strip() == "Montagem"
            ),
            None,
        )
        _assert(planning_row is not None, f"Planeamento nao mostrou backlog de montagem: {planning_rows}")
        _assert("falta stock" in str((planning_row or {}).get("obs", "") or "").lower(), f"Backlog de montagem sem alerta de stock: {planning_row}")

        finance = backend.finance_dashboard("Todos")
        status_map = {str(row.get("estado", "") or "").strip(): int(row.get("total", 0) or 0) for row in list(finance.get("order_status", []) or [])}
        _assert(status_map.get("Montagem", 0) >= 1, f"Dashboard financeiro sem estado Montagem: {finance}")
        finance_alert = next(
            (
                row
                for row in list(finance.get("montagem_alertas", []) or [])
                if str(row.get("numero", "") or "").strip() == assembly_only_order_num
            ),
            None,
        )
        _assert(finance_alert is not None, f"Dashboard financeiro sem alerta de montagem: {finance}")
        _assert("falta" in str((finance_alert or {}).get("stock", "") or "").lower(), f"Alerta de stock de montagem nao apareceu: {finance_alert}")
        _assert(float((finance_alert or {}).get("qtd_falta", 0) or 0) >= 3.0, f"Qtd em falta nao foi refletida no dashboard: {finance_alert}")
        _assert(
            str((finance_alert or {}).get("fornecedor", "") or "").strip() in acceptable_supplier_values,
            f"Fornecedor sugerido nao foi refletido: {finance_alert}",
        )

        purchase_result = backend.ne_create_from_montagem_shortages([assembly_only_order_num])
        purchase_note_num = str(purchase_result.get("numero", "") or "").strip()
        _assert(purchase_note_num, f"NE de montagem nao foi criada: {purchase_result}")
        _assert(not list(purchase_result.get("missing_supplier", []) or []), f"Fornecedor devia ter sido inferido: {purchase_result}")
        purchase_detail = backend.ne_detail(purchase_note_num)
        _assert(
            str(purchase_detail.get("fornecedor", "") or "").strip() in acceptable_supplier_values,
            f"Fornecedor da NE de montagem ficou errado: {purchase_detail}",
        )
        _assert(assembly_only_order_num in str(purchase_detail.get("obs", "") or ""), f"Origem da montagem nao ficou registada na obs: {purchase_detail}")
        purchase_line = next(
            (
                row
                for row in list(purchase_detail.get("lines", []) or [])
                if str(row.get("ref", "") or "").strip() == product_code
            ),
            None,
        )
        _assert(purchase_line is not None, f"Linha de compra da montagem nao foi criada: {purchase_detail}")
        _assert(abs(float((purchase_line or {}).get("qtd", 0) or 0) - 3.0) < 1e-6, f"Qtd da linha de compra devia refletir a falta real: {purchase_line}")
        _assert(
            str((purchase_line or {}).get("fornecedor_linha", "") or "").strip() in acceptable_supplier_values,
            f"Fornecedor sugerido da linha ficou errado: {purchase_line}",
        )

        print("conjuntos-montagem-flow-ok", quote_num, order_num, stock_before, stock_after)
        return 0
    finally:
        try:
            if purchase_note_num:
                backend.ne_remove(purchase_note_num)
        except Exception as exc:
            cleanup_errors.append(f"purchase_note_remove:{exc}")
        try:
            if purchase_history_num:
                backend.ne_remove(purchase_history_num)
        except Exception as exc:
            cleanup_errors.append(f"purchase_history_remove:{exc}")
        try:
            if assembly_only_order_num:
                backend.order_remove(assembly_only_order_num)
        except Exception as exc:
            cleanup_errors.append(f"assembly_order_remove:{exc}")
        try:
            if order_num:
                backend.order_remove(order_num)
        except Exception as exc:
            cleanup_errors.append(f"order_remove:{exc}")
        try:
            backend.orc_remove(assembly_only_quote_num)
        except Exception:
            pass
        try:
            backend.orc_remove(quote_num)
        except Exception:
            pass
        try:
            backend.assembly_model_remove(model_code)
        except Exception:
            pass
        try:
            backend.supplier_remove(supplier_id)
        except Exception as exc:
            cleanup_errors.append(f"supplier_remove:{exc}")
        try:
            data = backend.ensure_data()
            data["produtos"] = [
                row for row in list(data.get("produtos", []) or []) if str((row or {}).get("codigo", "") or "").strip() != product_code
            ]
            data["produtos_mov"] = [
                row
                for row in list(data.get("produtos_mov", []) or [])
                if str((row or {}).get("codigo", "") or "").strip() != product_code
                and str((row or {}).get("ref_doc", "") or "").strip() != product_code
            ]
            if created_client_code:
                data["clientes"] = [
                    row for row in list(data.get("clientes", []) or []) if str((row or {}).get("codigo", "") or "").strip() != created_client_code
                ]
            refs_db = data.setdefault("orc_refs", {})
            refs_db.pop(temp_ref_ext, None)
            data["seq"] = copy.deepcopy(seq_snapshot)
            data["of_seq"] = of_seq_snapshot
            data["opp_seq"] = opp_seq_snapshot
            data["orc_seq"] = orc_seq_snapshot
            backend._save(force=True)
        except Exception as exc:
            cleanup_errors.append(f"data_cleanup:{exc}")
        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except Exception:
            pass
        if cleanup_errors:
            raise RuntimeError("Falha na limpeza do teste: " + " | ".join(cleanup_errors))


if __name__ == "__main__":
    raise SystemExit(main())
