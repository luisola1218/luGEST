"""Material release stages stock, NC, notes and logs, including failed writes."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quality.application.material_release import MaterialRelease, ReleaseRules
from lugest_modules.quality.application.stock_policy import quarantine
from lugest_modules.quality.infrastructure.legacy_release_repository import LegacyReleaseRepository


def main():
    state = {"materiais": [{"id": "M1", "quantidade": 2, "quality_pending_qty": 3}],
             "quality_nonconformities": [{"id": "NC1", "material_id": "M1", "estado": "Aberta"}],
             "notas_encomenda": [{"numero": "NE1"}], "unrelated": [1]}
    original = deepcopy(state)
    fail = True
    saves = []
    def save(**kwargs):
        nonlocal state
        saves.append(kwargs)
        state = deepcopy(state)
        if fail:
            raise RuntimeError("write failed")
    def audit(data, **event):
        data.setdefault("audit_log", []).append(event)
    def log(data, action, details, **kwargs):
        data.setdefault("stock_log", []).append((action, details, kwargs))
    def sync(notes, materials):
        notes[0]["synced"] = materials[0]["quantidade"]
    now = lambda: "2026-09-10"
    parse = lambda value, default=0: float(value or default)
    repo = LegacyReleaseRepository(lambda: state, save, audit, log)
    service = MaterialRelease(repo, ReleaseRules(parse, now, lambda: "Tester",
        lambda item, **kw: quarantine(item, parse_float=parse, now_iso=now, **kw), sync))
    try:
        service.release("missing")
    except ValueError:
        pass
    else:
        raise AssertionError("missing NC accepted")
    assert state == original and not saves
    try:
        service.release("NC1")
    except RuntimeError:
        pass
    else:
        raise AssertionError("save failure swallowed")
    assert state == original
    fail = False
    assert service.release("NC1")["material_id"] == "M1"
    assert state["materiais"][0]["quantidade"] == 5
    assert state["materiais"][0]["quality_pending_qty"] == 0
    assert state["quality_nonconformities"][0]["estado"] == "Fechada"
    assert state["notas_encomenda"][0]["synced"] == 5
    assert len(state["stock_log"]) == 1
    service.release("NC1")
    assert state["materiais"][0]["quantidade"] == 5 and len(state["stock_log"]) == 1
    stale = repo.load()
    state["materiais"][0]["quantidade"] = 9
    try:
        repo.replace(stale, expected=stale, event={}, stock_event=None)
    except ValueError:
        pass
    else:
        raise AssertionError("stale snapshot accepted")
    assert all(call == dict(force=True, audit=False, blocking=True) for call in saves)
    print("quality-release-ok recovery=yes single-save=yes no-double-stock=yes stale=yes")

if __name__ == "__main__":
    main()
