"""Persist note fields absent from the legacy SQL table in its runtime payload."""
FIELDS = ("created_at", "data_aprovacao", "data_envio", "referencias_orcamento")
KEY = "purchase_note_metadata"


def export_metadata(notes):
    return {
        str(row.get("numero", "") or "").strip(): {
            key: str(row.get(key, "") or "").strip() for key in FIELDS if key in row
        }
        for row in notes or []
        if isinstance(row, dict) and str(row.get("numero", "") or "").strip()
    }


def apply_metadata(notes, payload):
    metadata = payload.get(KEY, {}) if isinstance(payload, dict) else {}
    if not isinstance(metadata, dict):
        return
    for row in notes or []:
        if not isinstance(row, dict):
            continue
        stored = metadata.get(str(row.get("numero", "") or "").strip())
        if isinstance(stored, dict):
            for key in FIELDS:
                if key in stored:
                    row[key] = str(stored[key] or "").strip()
