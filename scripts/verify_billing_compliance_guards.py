from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def main() -> int:
    backend = LegacyBackend()
    quote_num = ""
    order_num = ""
    record_num = ""
    try:
        quote = backend.orc_save(
            {
                "cliente": {"codigo": "CL0001", "nome": "Verto"},
                "estado": "Aprovado",
                "linhas": [
                    {
                        "tipo_item": "Peca",
                        "ref_externa": "VERIFY-COMPLIANCE-001",
                        "descricao": "VERIFY BILLING COMPLIANCE",
                        "material": "S235JR",
                        "espessura": "2",
                        "operacao": "Corte Laser + Embalamento",
                        "qtd": 5,
                        "preco_unit": 20.0,
                        "tempo_peca_min": 1,
                    }
                ],
            }
        )
        quote_num = str(quote.get("numero", "") or "").strip()
        backend.orc_set_state(quote_num, "Aprovado")
        converted = backend.orc_convert_to_order(quote_num, "VERIFY COMPLIANCE")
        order_num = str((converted.get("encomenda") or {}).get("numero", "") or "").strip()

        detail = backend.billing_open_record(source_type="quote", source_number=quote_num)
        record_num = str(detail.get("numero", "") or "").strip()
        if not record_num:
            raise RuntimeError("Registo de faturação sem número.")

        detail = backend.billing_add_invoice(
            record_num,
            {
                "numero_fatura": "FT-VERIFY-COMP-001",
                "serie": "2026A",
                "data_emissao": "2026-03-24",
                "data_vencimento": "2026-04-23",
                "valor_total": 123.00,
                "obs": "Teste de guardas de conformidade",
            },
        )
        invoice = dict((detail.get("invoices") or [{}])[0] or {})
        invoice_id = str(invoice.get("id", "") or "").strip()
        if not invoice_id:
            raise RuntimeError(f"Fatura não criada corretamente: {detail}")

        document = backend._billing_build_invoice_document(record_num, invoice)
        if str(document.get("software_cert", "") or "").strip():
            raise RuntimeError(f"O documento não devia trazer número AT hardcoded: {document.get('software_cert')!r}")
        if "0030/AT" in json.dumps(document, ensure_ascii=False):
            raise RuntimeError("Ainda existe referência hardcoded 0030/AT no documento de faturação.")

        detail = backend.billing_cancel_invoice(record_num, invoice_id, "Teste automático")
        cancelled = dict((detail.get("invoices") or [{}])[0] or {})
        if not bool(cancelled.get("anulada")):
            raise RuntimeError(f"Fatura devia ficar anulada: {cancelled}")
        if str(cancelled.get("estado", "") or "").strip() != "Anulada":
            raise RuntimeError(f"Estado da fatura anulada inválido: {cancelled}")

        row = next(
            row
            for row in backend.billing_rows("", "Todas", "Todos")
            if str(row.get("record_number", "") or "").strip() == record_num
        )
        if str(row.get("estado_faturacao", "") or "").strip() != "Por faturar":
            raise RuntimeError(f"Registo devia voltar a 'Por faturar' após anulação: {row}")
        if abs(float(row.get("faturado", 0) or 0) - 0.0) > 1e-6:
            raise RuntimeError(f"Valor faturado devia ignorar faturas anuladas: {row}")

        try:
            backend.billing_add_payment(
                record_num,
                {
                    "fatura_id": invoice_id,
                    "data_pagamento": "2026-03-24",
                    "valor": 50.0,
                    "metodo": "Transferência",
                },
            )
        except ValueError as exc:
            if "anulada" not in str(exc).lower():
                raise
        else:
            raise RuntimeError("Foi aceite um pagamento numa fatura anulada.")

        try:
            backend.billing_remove(record_num)
        except ValueError as exc:
            if "historico" not in str(exc).lower() and "histórico" not in str(exc).lower():
                raise
        else:
            raise RuntimeError("Foi permitido remover um registo de faturação com histórico documental.")

        print("billing-compliance-guards-ok", quote_num, order_num, record_num)
        return 0
    finally:
        data = backend.ensure_data()
        if record_num:
            data["faturacao"] = [
                row
                for row in list(data.get("faturacao", []) or [])
                if str((row or {}).get("numero", "") or "").strip() != record_num
            ]
        if order_num:
            try:
                backend.order_remove(order_num)
            except Exception:
                pass
        if quote_num:
            try:
                backend.orc_remove(quote_num)
            except Exception:
                pass
        backend._save(force=True)


if __name__ == "__main__":
    raise SystemExit(main())
