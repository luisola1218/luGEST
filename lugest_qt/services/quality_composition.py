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
