"""Explicit composition for migrated billing use cases."""
from uuid import uuid4
from lugest_modules.billing.application.payments import Payments, PaymentRules
from lugest_modules.billing.infrastructure.legacy_payment_repository import LegacyPaymentRepository


def payments(backend) -> Payments:
    return Payments(LegacyPaymentRepository(backend.ensure_data, backend._save),
                    PaymentRules(backend._parse_float, lambda: backend.desktop_main.now_iso(),
                                 lambda: uuid4().hex[:12].upper(), backend._billing_invoice_is_void))
