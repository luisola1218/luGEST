"""Nonconformity lifecycle, duplicate handling and audit recovery."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quality.application.nonconformities import Nonconformities, NonconformityRules
from lugest_modules.quality.infrastructure.legacy_nonconformity_repository import LegacyNonconformityRepository


def main():
    state = {}
    failure = False
    audit_failure = False

    def persist(**kwargs):
        nonlocal state
        assert kwargs == {"force": True, "audit": False, "blocking": True}
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")

    def audit(data, **event):
        data.setdefault("audit_log", []).append(deepcopy(event))
        if audit_failure:
            raise RuntimeError("audit failed")

    repository = LegacyNonconformityRepository(lambda: state, persist, audit)
    service = Nonconformities(repository, NonconformityRules(
        lambda value, default=0: float(str(value or default).replace(",", ".")),
        lambda: "2026-09-08T12:00:00", lambda: "TEST",
        lambda rows, prefix: f"{prefix}{len(rows)+1}", lambda kind, code: f"{kind}: {code}"))
    payload = {"origem": "Rececao fornecedor", "referencia": "NE-2026-0001 / lote",
               "entidade_tipo": "Material", "entidade_id": "M1", "descricao": "Rejeitado: 2,5"}
    created = service.save(payload)
    assert created["id"] == "NC1" and created["estado"] == "Aberta"
    assert service.rows()[0]["qtd_rejeitada"] == 2.5
    original = deepcopy(state)
    try:
        service.save(dict(payload, referencia="NE-2026-0001"))
    except ValueError:
        pass
    else:
        raise AssertionError("Open duplicate accepted")
    assert state == original
    service.close("NC1", "Verificado")
    assert service.rows() == [] and service.rows(state_filter="Fechadas")[0]["eficacia"] == "Verificado"
    assert state["audit_log"][-1]["action"] == "NC fechada"
    second = service.save(payload)
    assert second["id"] == "NC2"
    for fail_audit in (False, True):
        failure = not fail_audit
        audit_failure = fail_audit
        for action in (lambda: service.save(dict(payload, id="NC2", descricao="Alterada")),
                       lambda: service.close("NC2")):
            original = deepcopy(state)
            try:
                action()
            except RuntimeError:
                pass
            else:
                raise AssertionError("Failure hidden")
            assert state == original
    failure = audit_failure = False
    duplicate = deepcopy(second)
    duplicate["id"] = "NC3"
    state["quality_nonconformities"].append(duplicate)
    service.normalize_duplicates()
    assert state["quality_nonconformities"][-1]["estado"] == "Cancelada"
    stale = repository.rows()
    state["quality_nonconformities"][0]["descricao"] = "Concurrent"
    try:
        repository.replace([], expected=stale)
    except ValueError:
        pass
    else:
        raise AssertionError("Stale catalog accepted")
    assert "main" not in sys.modules
    print("quality-nc-ok lifecycle=yes duplicates=yes quantities=yes audit-restored=yes failures-restored=yes")


if __name__ == "__main__":
    main()
