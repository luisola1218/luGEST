"""Quality document metadata and audit recover together on immediate errors."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quality.application.documents import QualityDocuments, DocumentRules
from lugest_modules.quality.infrastructure.legacy_catalog_repository import LegacyQualityCatalogRepository


def main():
    state = {}
    failure = False
    audit_failure = False
    files = []
    def save(**kwargs):
        nonlocal state
        assert kwargs == dict(force=True, audit=False, blocking=True)
        state = deepcopy(state)
        if failure:
            raise RuntimeError("write failed")
    def audit(data, **event):
        data.setdefault("audit_log", []).append(event)
        if audit_failure:
            raise RuntimeError("audit failed")
    def store(source, title):
        files.append(source)
        return "stored/" + source
    repository = LegacyQualityCatalogRepository(lambda: state, save, audit, "quality_documents")
    service = QualityDocuments(repository, DocumentRules(lambda rows, prefix: "DOC1", store,
                                                         lambda: "2026-09-09T12:00:00", lambda: "TEST"))
    try:
        service.save({"caminho": "test.txt"})
    except ValueError:
        pass
    else:
        raise AssertionError("Missing title accepted")
    assert state == {} and files == []
    payload = {"titulo": "Certificado", "caminho": "test.txt", "entidade": "Material", "referencia": "M1"}
    row = service.save(payload)
    assert row["caminho"] == "stored/test.txt" and row["entidade_id"] == "M1"
    row["titulo"] = "Detached"
    assert service.rows("certificado")[0]["titulo"] == "Certificado"
    assert service.rows("missing") == []
    original = deepcopy(state)
    for audit_failure in (False, True):
        failure = not audit_failure
        for action in (lambda: service.save(dict(payload, id="DOC1", titulo="Edited", caminho="")),
                       lambda: service.remove("DOC1")):
            try:
                action()
            except RuntimeError:
                pass
            else:
                raise AssertionError("Failure hidden")
            assert state == original
    failure = audit_failure = False
    service.save(dict(payload, id="DOC1", titulo="Edited", caminho=""))
    assert service.rows()[0]["titulo"] == "Edited"
    service.remove("DOC1")
    assert not service.rows() and state["audit_log"][-1]["action"] == "Documento qualidade removido"
    original = deepcopy(state)
    try:
        service.remove("missing")
    except ValueError:
        pass
    else:
        raise AssertionError("Missing document removed")
    assert state == original and "main" not in sys.modules
    print("quality-documents-ok CRUD=yes detached=yes validation=yes audit-recovery=yes write-recovery=yes")


if __name__ == "__main__":
    main()
