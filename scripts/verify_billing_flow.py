from __future__ import annotations

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
                        "ref_externa": "VERIFY-BILL-001",
                        "descricao": "VERIFY BILLING FLOW",
                        "material": "S235JR",
                        "espessura": "2",
                        "operacao": "Corte Laser + Embalamento",
                        "qtd": 10,
                        "preco_unit": 12.5,
                        "tempo_peca_min": 1,
                    }
                ],
            }
        )
        quote_num = str(quote.get("numero", "") or "").strip()
        backend.orc_set_state(quote_num, "Aprovado")
        converted = backend.orc_convert_to_order(quote_num, "VERIFY BILLING")
        order_num = str((converted.get("encomenda") or {}).get("numero", "") or "").strip()

        detail = backend.billing_open_record(source_type="quote", source_number=quote_num)
        record_num = str(detail.get("numero", "") or "").strip()
        if not record_num:
            raise RuntimeError(f"Registo de faturação sem número: {detail}")

        detail = backend.billing_add_invoice(
            record_num,
            {
                "numero_fatura": "FT-VERIFY-001",
                "serie": "2026A",
                "data_emissao": "2026-03-23",
                "data_vencimento": "2026-04-22",
                "valor_total": 153.75,
                "obs": "Teste automático de faturação",
            },
        )
        invoice_id = str((detail.get("invoices") or [{}])[0].get("id", "") or "").strip()
        if not invoice_id:
            raise RuntimeError(f"Fatura não criada corretamente: {detail}")

        detail = backend.billing_add_payment(
            record_num,
            {
                "fatura_id": invoice_id,
                "data_pagamento": "2026-03-23",
                "valor": 153.75,
                "metodo": "Transferência",
                "referencia": "VERIFY-PAY",
                "titulo_comprovativo": "Comprovativo teste",
            },
        )

        row = next(
            row
            for row in backend.billing_rows("", "Todas", "Todos")
            if str(row.get("record_number", "") or "").strip() == record_num
        )
        if str(row.get("estado_faturacao", "") or "").strip() != "Faturada":
            raise RuntimeError(f"Estado de faturação inválido: {row}")
        if str(row.get("estado_pagamento", "") or "").strip() != "Paga":
            raise RuntimeError(f"Estado de pagamento inválido: {row}")
        if abs(float(row.get("saldo", 0) or 0) - 0.0) > 1e-6:
            raise RuntimeError(f"Saldo ainda aberto após pagamento total: {row}")

        dashboard = backend.billing_dashboard()
        if float(dashboard.get("received_total", 0) or 0) < 153.75:
            raise RuntimeError(f"Dashboard sem recebido esperado: {dashboard}")

        print("billing-flow-ok", quote_num, order_num, record_num)
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
