"""Payment edits, void invoice validation and write failure recovery."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.billing.application.payments import Payments, PaymentRules
from lugest_modules.billing.infrastructure.legacy_payment_repository import LegacyPaymentRepository


def main():
    state = {"faturacao": [{"numero": "R1", "faturas": [{"id": "I1"}], "pagamentos": []}]}
    failure = False
    saves = []

    def persist(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")

    repository = LegacyPaymentRepository(lambda: state, persist)
    service = Payments(repository, PaymentRules(lambda value, default=0: float(value or default),
                       lambda: "2026-09-08T12:00:00", lambda: "P1", lambda invoice: bool((invoice or {}).get("anulada"))))
    assert service.add("R1", {"valor": 20, "fatura_id": "I1"}) == "R1"
    service.add("R1", {"id": "P1", "valor": 25})
    assert state["faturacao"][0]["pagamentos"][0]["fatura_id"] == "I1"
    for value in (0, -1, float("nan"), float("inf")):
        original = deepcopy(state)
        try:
            service.add("R1", {"valor": value})
        except ValueError:
            pass
        else:
            raise AssertionError("Invalid payment accepted")
        assert state == original
    state["faturacao"][0]["faturas"][0]["anulada"] = True
    original = deepcopy(state)
    try:
        service.add("R1", {"id": "P1", "valor": 30})
    except ValueError:
        pass
    else:
        raise AssertionError("Omitted invoice ID bypassed void invoice validation")
    assert state == original
    state["faturacao"][0]["faturas"][0]["anulada"] = False
    failure = True
    for action in (lambda: service.add("R1", {"id": "P1", "valor": 40}),
                   lambda: service.add("R1", {"id": "P2", "valor": 5}),
                   lambda: service.remove("R1", "P1")):
        original = deepcopy(state)
        try:
            action()
        except RuntimeError:
            pass
        else:
            raise AssertionError("Payment write failure hidden")
        assert state == original
    failure = False
    stale = repository.get("R1")
    state["faturacao"][0]["obs"] = "Concurrent edit"
    try:
        repository.save(stale, expected=stale)
    except ValueError:
        pass
    else:
        raise AssertionError("Stale aggregate accepted")
    service.remove("R1", "P1")
    assert not state["faturacao"][0]["pagamentos"]
    assert all(call == {"force": True, "blocking": True} for call in saves)
    assert "main" not in sys.modules
    print("billing-payments-ok edit=yes void-invoice=yes finite-values=yes failures-restored=yes stale-rejected=yes")


if __name__ == "__main__":
    main()
