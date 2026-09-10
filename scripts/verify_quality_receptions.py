"""Reception decisions are isolated, movement-specific and recoverable."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quality.application.receptions import Receptions, ReceptionRules
from lugest_modules.quality.application.stock_policy import quarantine
from lugest_modules.quality.infrastructure.legacy_reception_repository import LegacyReceptionRepository


def main():
    def note(number):
        return {"numero": number, "linhas": [{"origem": "Produto", "ref": "P1", "entregas_linha": [
            {"qtd": 5, "quality_status": "EM_INSPECAO", "quality_approved_qty": 0, "quality_rejected_qty": 0}]}]}
    state = {"produtos": [{"codigo": "P1", "qty": 0}], "notas_encomenda": [note("NE1"), note("NE2")],
             "materiais": [], "quality_nonconformities": [], "quality_documents": [], "unrelated": [1]}
    saves = []
    fail = False
    def save(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)
        if fail:
            raise RuntimeError("write failed")
    def audit(data, **event):
        data.setdefault("audit_log", []).append(event)
    def stock(data, action, details, **kwargs):
        data.setdefault("stock_log", []).append((action, details))
    def product(data, **event):
        data.setdefault("produtos_mov", []).append(event)
    parse = lambda value, default=0: default if value is None else float(value)
    now = lambda: "2026-09-10"
    rules = ReceptionRules(parse, now, lambda: "Tester", lambda rows, prefix: prefix + str(len(rows)+1),
                           str, lambda value: value == "Material",
                           lambda item, **kw: quarantine(item, parse_float=parse, now_iso=now, **kw),
                           lambda notes, materials: None)
    repo = LegacyReceptionRepository(lambda: state, save, audit, stock, product)
    service = Receptions(repo, rules)
    original = deepcopy(state)
    rows = service.rows()
    assert len(rows) == 2 and state == original and not saves
    payload = dict(tipo="Produto", id="P1", movement_id=rows[0]["movement_id"], quality_status="APROVADO")
    for value in (6, float("nan"), float("inf")):
        try:
            service.save(dict(payload, qtd_aprovada=value))
        except ValueError:
            pass
        else:
            raise AssertionError("invalid quantity accepted")
        assert state == original and not saves
    fail = True
    try:
        service.save(payload)
    except RuntimeError:
        pass
    else:
        raise AssertionError("save failure swallowed")
    assert state == original
    fail = False
    other = deepcopy(state["notas_encomenda"][1])
    service.save(dict(payload, qtd_aprovada=2))
    assert state["produtos"][0]["qty"] == 2
    assert state["notas_encomenda"][1] == other
    assert state["produtos"][0]["quality_pending_qty"] == 8
    service.save(dict(payload, quality_status="DEVOLVER_FORNECEDOR", qtd_rejeitada=1, qtd_aprovada=0))
    assert state["quality_documents"][0]["qtd"] == 1
    assert state["quality_nonconformities"][0]["qtd_rejeitada"] == 1
    assert state["produtos"][0]["quality_pending_qty"] == 7
    service.save(payload)
    assert state["produtos"][0]["qty"] == 4
    assert state["produtos"][0]["quality_pending_qty"] == 5
    assert state["notas_encomenda"][1] == other
    service.save(payload)
    assert state["produtos"][0]["qty"] == 4
    before = deepcopy(state)
    service.reconcile()
    assert state == before
    assert all(call == dict(force=True, audit=False, blocking=True) for call in saves)
    assert len(state["produtos_mov"]) == 2
    from lugest_modules.quality.infrastructure.reception_metadata import export_metadata, apply_metadata, KEY
    metadata = export_metadata(state)
    restored = deepcopy(state)
    for note in restored["notas_encomenda"]:
        for line in note["linhas"]:
            for movement in line["entregas_linha"]:
                for field in ("quality_approved_qty", "quality_rejected_qty", "quality_pending_qty", "quality_movement_id"):
                    movement.pop(field, None)
    restored["quality_documents"][0].pop("qtd")
    apply_metadata(restored, {KEY: metadata})
    assert restored == state
    changed = deepcopy(restored)
    movement = changed["notas_encomenda"][0]["linhas"][0]["entregas_linha"][0]
    movement["stock_ref"] = "ANOTHER"
    movement.pop("quality_approved_qty")
    apply_metadata(changed, {KEY: metadata})
    assert "quality_approved_qty" not in movement
    print("quality-receptions-ok partial=yes rejected=yes returns=yes isolation=yes failures=yes finite=yes")

if __name__ == "__main__":
    main()
