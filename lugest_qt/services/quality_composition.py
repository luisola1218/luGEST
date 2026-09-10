"""Explicit dependencies for the quality business module."""
from datetime import datetime
from lugest_modules.quality.application.nonconformities import Nonconformities, NonconformityRules
from lugest_modules.quality.infrastructure.legacy_nonconformity_repository import LegacyNonconformityRepository


def nonconformities(backend) -> Nonconformities:
    return Nonconformities(
        LegacyNonconformityRepository(backend.ensure_data, backend._save, backend._append_audit_event),
        NonconformityRules(backend._parse_float,
                          lambda: backend.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"),
                          backend._current_user_label, backend._next_prefixed_id,
                          lambda *args: backend._quality_link_label(*args)),
    )


from lugest_modules.quality.application.documents import QualityDocuments, DocumentRules
from lugest_modules.quality.infrastructure.legacy_catalog_repository import LegacyQualityCatalogRepository


def quality_documents(backend) -> QualityDocuments:
    return QualityDocuments(
        LegacyQualityCatalogRepository(backend.ensure_data, backend._save, backend._append_audit_event, "quality_documents"),
        DocumentRules(backend._next_prefixed_id,
                      lambda source, title: backend._store_shared_file(source, "quality/documents", preferred_name=backend._file_reference_name(source, title)),
                      lambda: backend.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"),
                      backend._current_user_label),
    )


from lugest_modules.quality.application.material_release import MaterialRelease, ReleaseRules
from lugest_modules.quality.application.stock_policy import quarantine
from lugest_modules.quality.infrastructure.legacy_release_repository import LegacyReleaseRepository
from lugest_modules.purchasing.application.material_lines import sync_material_lines
from lugest_qt.services.purchasing_composition import material_line_rules


def material_release(backend) -> MaterialRelease:
    now = lambda: backend.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds")
    def sync_notes(notes, materials):
        rules = material_line_rules(backend)
        for note in notes:
            sync_material_lines(note, materials, rules)
    return MaterialRelease(
        LegacyReleaseRepository(backend.ensure_data, backend._save, backend._append_audit_event,
                                backend.desktop_main.log_stock),
        ReleaseRules(backend._parse_float, now, backend._current_user_label,
                     lambda item, **kwargs: quarantine(item, parse_float=backend._parse_float, now_iso=now, **kwargs),
                     sync_notes),
    )
