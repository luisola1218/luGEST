"""Round-trip quality fields absent from relational tables in the existing runtime payload.

Movement metadata is scoped by note and position and checked against stock reference
and quantity. It never creates rows or overwrites relational identity/status fields.
"""
from copy import deepcopy

KEY = "quality_reception_metadata"
MOVEMENT_FIELDS = ("quality_movement_id", "quality_approved_qty", "quality_rejected_qty", "quality_pending_qty", "inspection_at", "inspection_by")
ITEM_FIELDS = ("quality_rejected_qty", "quality_last_received_qty", "quality_last_approved_qty", "quality_last_rejected_qty", "quality_released_at", "quality_released_by")
DOCUMENT_FIELDS = ("nc_id", "qtd", "qtd_stock")


def selected(row, fields):
    return {key: deepcopy(row[key]) for key in fields if key in row}


def movements(notes):
    for note in notes or []:
        for index, line in enumerate(note.get("linhas", []) or []):
            for offset, movement in enumerate(line.get("entregas_linha", []) or []):
                yield f"{note.get('numero', '')}|{index}|{offset}", line, movement


def identity(line, movement):
    return [str(movement.get("stock_ref") or line.get("ref", "")).strip(), float(movement.get("qtd", 0) or 0)]


def export_metadata(data):
    result = {"movements": {key: {"identity": identity(line, row), "fields": selected(row, MOVEMENT_FIELDS)}
                            for key, line, row in movements(data.get("notas_encomenda", []))}}
    for bucket, identifier, fields in (("materiais", "id", ITEM_FIELDS), ("produtos", "codigo", ITEM_FIELDS),
                                       ("quality_documents", "id", DOCUMENT_FIELDS)):
        result[bucket] = {str(row.get(identifier, "")).strip(): selected(row, fields) for row in data.get(bucket, []) or []}
    return result


def apply_metadata(data, payload):
    stored = payload.get(KEY, {}) if isinstance(payload, dict) else {}
    if not isinstance(stored, dict):
        return
    movement_data = stored.get("movements", {})
    if isinstance(movement_data, dict):
        for key, line, row in movements(data.get("notas_encomenda", [])):
            entry = movement_data.get(key, {})
            if isinstance(entry, dict) and entry.get("identity") == identity(line, row) and isinstance(entry.get("fields"), dict):
                row.update(selected(entry["fields"], MOVEMENT_FIELDS))
    for bucket, identifier, fields in (("materiais", "id", ITEM_FIELDS), ("produtos", "codigo", ITEM_FIELDS),
                                       ("quality_documents", "id", DOCUMENT_FIELDS)):
        values = stored.get(bucket, {})
        if not isinstance(values, dict):
            continue
        for row in data.get(bucket, []) or []:
            entry = values.get(str(row.get(identifier, "")).strip())
            if isinstance(entry, dict):
                row.update(selected(entry, fields))
