"""Purchase transitions and batch conversion must not publish partial writes."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.purchasing.application.note_status import normalize_status
from lugest_modules.purchasing.infrastructure.note_metadata import KEY, export_metadata, apply_metadata
from lugest_modules.purchasing.application.note_lifecycle import NoteLifecycle, NoteRules
from lugest_modules.purchasing.infrastructure.legacy_note_repository import LegacyNoteRepository


def main():
    state = {}
    failure = False
    allocations = []
    def allocate(candidate):
        value = candidate.setdefault("seq", {}).get("ne", 1)
        candidate["seq"]["ne"] = value + 1
        allocations.append(value)
        return f"NE{value}"
    def save(**kwargs):
        nonlocal state
        assert kwargs == dict(force=True, blocking=True)
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")
    repository = LegacyNoteRepository(lambda: state, save, allocate)
    service = NoteLifecycle(repository, NoteRules(
        lambda: "2026-09-09T12:00:00", lambda note: "rfq",
        lambda name: (name if name in {"F1", "F2"} else "", name, ""),
        lambda note: note.update(total=sum(row.get("total", 0) for row in note["linhas"])),
    ))
    draft = service.create_draft()
    assert draft["numero"] == "NE1"
    draft["obs"] = "Detached"
    assert repository.rows()[0]["obs"] == ""
    original = deepcopy(state)
    try:
        service.approve("NE1")
    except ValueError:
        pass
    else:
        raise AssertionError("Empty note approved")
    assert state == original
    state["notas_encomenda"][0]["linhas"] = [
        {"ref": "P1", "fornecedor_linha": "F1", "total": 10, "entregas_linha": [{"qtd": 1}]},
        {"ref": "P2", "fornecedor_linha": "UNKNOWN", "total": 20},
    ]
    original = deepcopy(state)
    count = len(allocations)
    try:
        service.generate_orders("NE1")
    except ValueError:
        pass
    else:
        raise AssertionError("Unregistered supplier accepted")
    assert state == original and len(allocations) == count
    state["notas_encomenda"][0]["linhas"][1]["fornecedor_linha"] = "F2"
    original = deepcopy(state)
    failure = True
    for action in (lambda: service.create_draft(), lambda: service.approve("NE1"),
                   lambda: service.mark_sent("NE1"), lambda: service.remove("NE1"),
                   lambda: service.generate_orders("NE1")):
        try:
            action()
        except RuntimeError:
            pass
        else:
            raise AssertionError("Persistence failure hidden")
        assert state == original
    failure = False
    assert service.approve("NE1")["estado"] == "Cotacao aprovada"
    assert service.mark_sent("NE1")["estado"] == "Enviada"
    created = service.generate_orders("NE1")
    assert len(created) == 2 and [row["total"] for row in created] == [10, 20]
    assert repository.rows()[0]["ne_geradas"] == [row["numero"] for row in created]
    assert "entregas_linha" not in repository.rows()[1]["linhas"][0]
    original = deepcopy(state)
    try:
        service.generate_orders("NE1")
    except ValueError:
        pass
    else:
        raise AssertionError("Repeated conversion duplicated orders")
    assert state == original
    stale = repository.rows()
    state["notas_encomenda"][0]["obs"] = "Concurrent"
    try:
        repository.replace([], expected=stale)
    except ValueError:
        pass
    else:
        raise AssertionError("Stale note catalog accepted")
    service.remove(created[0]["numero"])
    assert len(repository.rows()) == 2 and "main" not in sys.modules
    notes = [{"numero": "NE1", "created_at": "2026-09-09", "data_aprovacao": "2026-09-10T08:00:00", "data_envio": "2026-09-10T09:00:00", "referencias_orcamento": "ORC1"}]
    metadata = export_metadata(notes)
    restored = [{"numero": "NE1", "estado": "Enviada"}]
    apply_metadata(restored, {KEY: metadata})
    assert all(restored[0][key] == value for key, value in notes[0].items())
    apply_metadata(restored, {})
    assert restored[0]["data_aprovacao"] == notes[0]["data_aprovacao"]
    assert "NE1" not in export_metadata([])
    for status in ("Enviada", "Cotacao aprovada", "Aprovada", "Em edicao"):
        note = {"estado": status, "linhas": [{"qtd": 2}]}
        normalize_status(note, lambda value, default=0: float(value or default), str.lower)
        assert note["estado"] == status
    note = {"estado": "Enviada", "linhas": [{"qtd": 2, "qtd_entregue": 1}]}
    normalize_status(note, lambda value, default=0: float(value or default), str.lower)
    assert note["estado"] == "Parcialmente Entregue"
    note["linhas"][0]["qtd_entregue"] = 2
    normalize_status(note, lambda value, default=0: float(value or default), str.lower)
    assert note["estado"] == "Entregue"
    print("purchase-lifecycle-ok drafts=yes transitions=yes conversion=yes validation=yes recovery=yes stale=yes")


if __name__ == "__main__":
    main()
