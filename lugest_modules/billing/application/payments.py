"""Billing payment commands with detached aggregates and explicit persistence."""
from copy import deepcopy
from dataclasses import dataclass
from math import isfinite
from typing import Any, Callable, Protocol

class PaymentRepository(Protocol):
    def get(self, number: str) -> dict[str, Any] | None: ...
    def save(self, record: dict[str, Any], *, expected: dict[str, Any]) -> None: ...

@dataclass(frozen=True)
class PaymentRules:
    parse_float: Callable
    now_iso: Callable
    new_id: Callable
    invoice_is_void: Callable

class Payments:
    def __init__(self, repository: PaymentRepository, rules: PaymentRules):
        self.repository = repository
        self.rules = rules

    def normalize(self, payload: dict[str, Any], existing: dict[str, Any] | None = None) -> dict[str, Any]:
        row = deepcopy(existing or {})
        row_id = str(payload.get("id", "") or row.get("id", "") or self.rules.new_id()).strip()
        data_pagamento = str(payload.get("data_pagamento", "") or row.get("data_pagamento", "") or "").strip()[:10]
        valor = round(self.rules.parse_float(payload.get("valor", row.get("valor", 0)), 0), 2)
        metodo = str(payload.get("metodo", "") or row.get("metodo", "") or "").strip()
        referencia = str(payload.get("referencia", "") or row.get("referencia", "") or "").strip()
        titulo = str(payload.get("titulo_comprovativo", "") or row.get("titulo_comprovativo", "") or "").strip()
        caminho = str(payload.get("caminho_comprovativo", "") or row.get("caminho_comprovativo", "") or "").strip()
        fatura_id = str(payload.get("fatura_id", "") or row.get("fatura_id", "") or "").strip()
        obs = str(payload.get("obs", "") or row.get("obs", "") or "").strip()
        if not isfinite(valor) or valor <= 0:
            raise ValueError("Valor do pagamento invalido.")
        if not data_pagamento:
            data_pagamento = str(self.rules.now_iso())[:10]
        return {
            "id": row_id,
            "fatura_id": fatura_id,
            "data_pagamento": data_pagamento,
            "valor": valor,
            "metodo": metodo,
            "referencia": referencia,
            "titulo_comprovativo": titulo,
            "caminho_comprovativo": caminho,
            "obs": obs,
            "created_at": str(row.get("created_at", "") or self.rules.now_iso()),
        }


    def add(self, numero: str, payload: dict[str, Any]) -> str:
        reg_num = str(numero or "").strip()
        record = self.repository.get(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        original = deepcopy(record)
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        existing = next((row for row in list(record.get("pagamentos", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        invoice_id = str(payload_dict.get("fatura_id", "") or (existing or {}).get("fatura_id", "") or "").strip()
        if invoice_id and not any(str(row.get("id", "") or "").strip() == invoice_id for row in list(record.get("faturas", []) or [])):
            raise ValueError("A fatura associada ao pagamento não existe neste registo.")
        if invoice_id:
            invoice = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == invoice_id), None)
            if self.rules.invoice_is_void(invoice):
                raise ValueError("Nao e possivel associar pagamentos a uma fatura anulada.")
        payment = self.normalize(payload_dict, existing)
        if existing is None:
            record.setdefault("pagamentos", []).append(payment)
        else:
            existing.update(payment)
        record["updated_at"] = self.rules.now_iso()
        self.repository.save(record, expected=original)
        return reg_num


    def remove(self, numero: str, payment_id: str) -> str:
        reg_num = str(numero or "").strip()
        row_id = str(payment_id or "").strip()
        record = self.repository.get(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        original = deepcopy(record)
        before = len(list(record.get("pagamentos", []) or []))
        record["pagamentos"] = [row for row in list(record.get("pagamentos", []) or []) if str(row.get("id", "") or "").strip() != row_id]
        if len(record["pagamentos"]) == before:
            raise ValueError("Pagamento não encontrado.")
        record["updated_at"] = self.rules.now_iso()
        self.repository.save(record, expected=original)
        return reg_num


