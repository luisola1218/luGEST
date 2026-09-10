"""Connect purchase use cases to historical persistence and value rules."""
from lugest_modules.purchasing.application.note_lifecycle import NoteLifecycle, NoteRules
from lugest_modules.purchasing.infrastructure.legacy_note_repository import LegacyNoteRepository


def note_lifecycle(backend) -> NoteLifecycle:
    return NoteLifecycle(
        LegacyNoteRepository(backend.ensure_data, backend._save, backend.desktop_main.next_ne_numero),
        NoteRules(backend.desktop_main.now_iso, backend._note_kind,
                  backend._resolve_supplier, backend._recalculate_note_totals),
    )
