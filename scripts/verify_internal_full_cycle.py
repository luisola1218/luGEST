from __future__ import annotations

from datetime import date, timedelta
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_infra.storage import files as lugest_storage
from lugest_qt.services.main_bridge import LegacyBackend


def _reset_piece_ops(backend: LegacyBackend, enc_num: str, piece_id: str) -> None:
    reset_fn = getattr(backend.operador_actions, "_mysql_ops_reset_piece", None)
    if callable(reset_fn):
        reset_fn(enc_num, piece_id)


def _start_with_retry(backend: LegacyBackend, enc_num: str, piece_id: str, operator_name: str, operation: str) -> None:
    try:
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")
    except ValueError as exc:
        if "Operacao ocupada" not in str(exc):
            raise
        _reset_piece_ops(backend, enc_num, piece_id)
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")


def _finish_with_retry(
    backend: LegacyBackend,
    enc_num: str,
    piece_id: str,
    operator_name: str,
    operation: str,
    ok_qty: float,
) -> None:
    last_exc = None
    for _attempt in range(3):
        try:
            backend.operator_finish_piece(enc_num, piece_id, operator_name, ok_qty, 0, 0, operation, "Geral")
            return
        except ValueError as exc:
            message = str(exc)
            if ("Estado atual: Livre" not in message) and ("Inicia primeiro a operacao" not in message):
                raise
            last_exc = exc
            _reset_piece_ops(backend, enc_num, piece_id)
            _start_with_retry(backend, enc_num, piece_id, operator_name, operation)
    if last_exc is not None:
        raise last_exc


def _find_piece_id(backend: LegacyBackend, enc_num: str) -> tuple[str, dict]:
    detail = backend.order_detail(enc_num)
    pieces = list(detail.get("pieces", []) or [])
    if not pieces:
        raise RuntimeError(f"Encomenda sem peças para operar: {enc_num}")
    piece = pieces[0]
    return str(piece.get("id", "") or "").strip(), piece


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    clients = list(data.get("clientes", []) or [])
    materials = list(data.get("materiais", []) or [])
    products = list(data.get("produtos", []) or [])
    refs_count = len(dict(data.get("orc_refs", {}) or {}))
    if len(clients) < 5:
        raise RuntimeError(f"Esperados pelo menos 5 clientes na base de teste, obtidos {len(clients)}.")
    if len(materials) < 10:
        raise RuntimeError(f"Esperados pelo menos 10 materiais na base de teste, obtidos {len(materials)}.")
    if len(products) < 10:
        raise RuntimeError(f"Esperados pelo menos 10 produtos na base de teste, obtidos {len(products)}.")
    if refs_count < 25:
        raise RuntimeError(f"Esperadas pelo menos 25 referências históricas, obtidas {refs_count}.")

    counts = {str(row.get("title", "") or ""): str(row.get("value", "") or "") for row in backend.dashboard_counts()}
    if counts.get("Clientes") != str(len(clients)):
        raise RuntimeError(f"Dashboard geral de clientes incorreto: {counts}")
    if counts.get("Materias") != str(len(materials)):
        raise RuntimeError(f"Dashboard geral de materiais incorreto: {counts}")

    today_iso = str(backend.desktop_main.now_iso())[:10]
    due_iso = (date.fromisoformat(today_iso) + timedelta(days=30)).isoformat()
    client = dict(clients[0])
    material = dict(materials[0])

    quote = backend.orc_save(
        {
            "cliente": {"codigo": client["codigo"], "nome": client["nome"]},
            "estado": "Aprovado",
            "linhas": [
                {
                    "tipo_item": "Peca",
                    "ref_externa": f"E2E-{client['codigo']}-001",
                    "descricao": "Fluxo interno completo faturação",
                    "material": str(material.get("material", "") or ""),
                    "espessura": str(material.get("espessura", "") or ""),
                    "operacao": "Corte Laser + Embalamento",
                    "qtd": 12,
                    "preco_unit": 24.80,
                    "tempo_peca_min": 3,
                }
            ],
            "iva_perc": 23,
            "preco_transporte": 0.0,
            "nota_cliente": "Exercício Orçamento -> Faturação",
            "executado_por": "admin",
        }
    )
    quote_num = str(quote.get("numero", "") or "").strip()
    backend.orc_set_state(quote_num, "Aprovado")
    converted = backend.orc_convert_to_order(quote_num, "Fluxo completo interno")
    order_num = str((converted.get("encomenda") or {}).get("numero", "") or "").strip()
    if not order_num:
        raise RuntimeError(f"Falha ao converter orçamento em encomenda: {converted}")

    pending_rows = [row for row in backend.planning_pending_rows(order_num, "Pendentes") if str(row.get("numero", "") or "").strip() == order_num]
    if not pending_rows:
        raise RuntimeError(f"Planeamento não recebeu a encomenda criada: {order_num}")
    placed = backend.planning_auto_plan(pending_rows, week_start=today_iso)
    if not placed:
        raise RuntimeError("Auto planeamento não criou blocos para a encomenda de teste.")

    piece_id, piece = _find_piece_id(backend, order_num)
    _reset_piece_ops(backend, order_num, piece_id)
    _start_with_retry(backend, order_num, piece_id, "admin", "Corte Laser")
    _finish_with_retry(backend, order_num, piece_id, "admin", "Corte Laser", 12)
    _start_with_retry(backend, order_num, piece_id, "admin", "Embalamento")
    _finish_with_retry(backend, order_num, piece_id, "admin", "Embalamento", 12)

    available = backend.expedicao_available_pieces(order_num)
    if len(available) != 1:
        raise RuntimeError(f"Expedição sem disponibilidade correta: {available}")
    if abs(float(available[0].get("disponivel_num", 0) or 0) - 12.0) > 1e-6:
        raise RuntimeError(f"Quantidade disponível para guia incorreta: {available}")

    line = dict(available[0])
    line["peca_id"] = piece_id
    line["qtd"] = 12.0
    line["peso"] = 0.0
    line["unid"] = "UN"
    guide = backend.expedicao_emit_off(order_num, [line], backend.expedicao_defaults_for_order(order_num))
    guide_num = str(guide.get("numero", "") or "").strip()
    if not guide_num:
        raise RuntimeError(f"Guia não foi emitida: {guide}")
    if not any(str(row.get("numero", "") or "").strip() == guide_num for row in backend.expedicao_rows(order_num)):
        raise RuntimeError(f"Guia não apareceu no histórico da encomenda: {guide_num}")

    record = backend.billing_open_record(source_type="quote", source_number=quote_num)
    record_num = str(record.get("numero", "") or "").strip()
    if not record_num:
        raise RuntimeError(f"Registo de faturação inválido: {record}")
    invoice_payload = backend.billing_invoice_defaults(record_num)
    invoice_payload["guia_numero"] = guide_num
    invoice_payload["data_emissao"] = today_iso
    invoice_payload["data_vencimento"] = due_iso
    detail = backend.billing_generate_invoice_pdf(record_num, invoice_payload)
    invoices = list(detail.get("invoices", []) or [])
    if len(invoices) != 1:
        raise RuntimeError(f"Fatura não foi criada corretamente: {detail}")
    invoice = dict(invoices[0])
    invoice_num = str(invoice.get("numero_fatura", "") or "").strip()
    pdf_path = lugest_storage.resolve_file_reference(str(invoice.get("caminho", "") or "").strip(), base_dir=ROOT)
    if not invoice_num or pdf_path is None or not pdf_path.exists():
        raise RuntimeError(f"PDF de fatura não ficou disponível: {invoice}")
    if pdf_path.stat().st_size < 2500:
        raise RuntimeError(f"PDF de fatura demasiado pequeno, provável render incompleta: {pdf_path}")
    pdf_text = pdf_path.read_text(encoding="latin-1", errors="ignore")
    for token in (invoice_num, guide_num, str(client.get("nome", "") or "").split()[0], "Fatura", "Base tributavel"):
        if token and token not in pdf_text:
            raise RuntimeError(f"PDF da fatura não contém o elemento esperado `{token}`.")

    proof_dir = ROOT / "generated" / "comprovativos"
    proof_dir.mkdir(parents=True, exist_ok=True)
    proof_path = proof_dir / f"{invoice_num}_pagamento.txt"
    proof_path.write_text(
        "\n".join(
            [
                f"Comprovativo interno de pagamento {invoice_num}",
                f"Data: {today_iso}",
                f"Valor: {float(invoice.get('valor_total', 0) or 0):.2f} EUR",
                f"Referencia: PAG-{record_num}",
            ]
        ),
        encoding="utf-8",
    )
    detail = backend.billing_add_payment(
        record_num,
        {
            "fatura_id": str(invoice.get("id", "") or "").strip(),
            "data_pagamento": today_iso,
            "valor": float(invoice.get("valor_total", 0) or 0),
            "metodo": "Transferência",
            "referencia": f"PAG-{record_num}",
            "titulo_comprovativo": f"Pagamento {invoice_num}",
            "caminho_comprovativo": str(proof_path),
            "obs": "Pagamento integral do exercício interno",
        },
    )

    billing_row = next(
        row for row in backend.billing_rows("", "Todas", "Todos") if str(row.get("record_number", "") or "").strip() == record_num
    )
    if str(billing_row.get("estado_faturacao", "") or "").strip() != "Faturada":
        raise RuntimeError(f"Estado de faturação incorreto no dashboard: {billing_row}")
    if str(billing_row.get("estado_pagamento", "") or "").strip() != "Paga":
        raise RuntimeError(f"Estado de pagamento incorreto no dashboard: {billing_row}")
    if abs(float(billing_row.get("saldo", 0) or 0) - 0.0) > 1e-6:
        raise RuntimeError(f"Saldo do registo não ficou liquidado: {billing_row}")

    billing_dashboard = backend.billing_dashboard()
    if float(billing_dashboard.get("invoiced_total", 0) or 0) < float(invoice.get("valor_total", 0) or 0):
        raise RuntimeError(f"Dashboard de faturação sem total faturado esperado: {billing_dashboard}")
    if float(billing_dashboard.get("received_total", 0) or 0) < float(invoice.get("valor_total", 0) or 0):
        raise RuntimeError(f"Dashboard de faturação sem total recebido esperado: {billing_dashboard}")

    final_order = backend.get_encomenda_by_numero(order_num) or {}
    shipping_state = str(final_order.get("estado_expedicao", "") or "").strip()
    if "exped" not in shipping_state.lower():
        raise RuntimeError(f"Estado de expedição da encomenda não ficou coerente: {shipping_state}")

    print(
        json.dumps(
            {
                "quote": quote_num,
                "order": order_num,
                "guide": guide_num,
                "billing_record": record_num,
                "invoice": invoice_num,
                "invoice_pdf": str(pdf_path),
                "payment_proof": str(proof_path),
                "billing_dashboard": billing_dashboard,
                "counts": counts,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
