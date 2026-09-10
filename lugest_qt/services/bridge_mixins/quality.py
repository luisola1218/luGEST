from __future__ import annotations
from lugest_qt.services.quality_composition import nonconformities, quality_documents, material_release, receptions, quality_queries

from lugest_modules.quality.application.stock_policy import status_code, quarantine

from datetime import datetime
from typing import Any


class QualityBackendMixin:
    """Legacy adapter for quality; see BACKEND_GUIDE.md."""

    def quality_summary(self) -> dict[str, Any]:
        return quality_queries(self).summary()

    def quality_data_health(self) -> dict[str, Any]:
        return quality_queries(self).health()

    def quality_link_options(self) -> dict[str, list[dict[str, str]]]:
        return quality_queries(self).link_options()

    def _quality_link_label(self, entity_type: str, entity_id: str) -> str:
        return quality_queries(self).link_label(entity_type, entity_id)

    def _quality_status_code(self, value: Any) -> str:
        return status_code(value)

    def _quality_status_is_available(self, value: Any) -> bool:
        return status_code(value) == "APROVADO"

    def _quality_quarantine_pending_stock(self, item: dict[str, Any], *, kind: str, max_qty: float | None = None) -> bool:
        return quarantine(item, kind=kind, max_qty=max_qty, parse_float=self._parse_float,
                          now_iso=lambda: self.desktop_main.now_iso() or datetime.now().isoformat(timespec="seconds"))

    def _quality_movement_pending_qty(self, movement: dict[str, Any]) -> float:
        qty = self._parse_float(movement.get("qtd", 0), 0)
        approved = self._parse_float(movement.get("quality_approved_qty", 0), 0)
        rejected = self._parse_float(movement.get("quality_rejected_qty", 0), 0)
        status = self._quality_status_code(movement.get("quality_status", movement.get("inspection_status", "")))
        if approved <= 0 and rejected <= 0:
            if status == "APROVADO":
                approved = qty
            elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                rejected = qty
        return max(0.0, qty - approved - rejected)

    def _quality_delivery_movement_id(self, note_number: str, line_index: int, movement_index: int, entity_type: str, entity_id: str) -> str:
        return "|".join(
            str(part or "").strip()
            for part in (
                note_number,
                f"L{int(line_index or 0)}",
                f"M{int(movement_index or 0)}",
                entity_type,
                entity_id,
            )
        )

    def _quality_reception_entity_keys(self) -> set[tuple[str, str]]:
        return {
            (str(row.get("entity_type", "") or "").strip(), str(row.get("entity_id", "") or "").strip())
            for row in self._quality_iter_delivery_movements(ensure_ids=False)
            if str(row.get("entity_type", "") or "").strip() and str(row.get("entity_id", "") or "").strip()
        }

    def _quality_iter_delivery_movements(self, *, ensure_ids: bool = False) -> list[dict[str, Any]]:
        return receptions(self).movement_rows()

    def _quality_sync_pending_from_delivery_movements(self) -> None:
        receptions(self).reconcile()

    def _quality_reference_key(self, value: Any, fallback: Any = "") -> str:
        return nonconformities(self).reference_key(value, fallback)

    def _quality_nc_key(self, payload: dict[str, Any]) -> tuple[str, str, str, str]:
        return nonconformities(self).key(payload)

    def _quality_is_open_nc(self, row: dict[str, Any]) -> bool:
        return nonconformities(self).is_open(row)

    def _quality_find_open_nc(self, payload: dict[str, Any], *, exclude_id: str = "") -> dict[str, Any] | None:
        return nonconformities(self).find_open(payload, exclude_id=exclude_id)

    def _quality_nc_quantity(self, row: dict[str, Any] | None, field: str) -> float:
        return nonconformities(self).quantity(row, field)

    def _quality_normalize_open_nc_duplicates(self) -> None:
        return nonconformities(self).normalize_duplicates()

    def quality_reception_rows(self, filter_text: str = "", state_filter: str = "Pendentes") -> list[dict[str, Any]]:
        return receptions(self).rows(filter_text, state_filter)

    def quality_reception_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return receptions(self).save(payload)

    def _quality_return_document(
        self,
        target: dict[str, Any],
        *,
        entity_type: str,
        entity_id: str,
        reference: str,
        nc_id: str = "",
    ) -> dict[str, Any] | None:
        return receptions(self).return_document(target, entity_type=entity_type, entity_id=entity_id, reference=reference, nc_id=nc_id)

    def quality_nc_rows(self, filter_text: str = "", state_filter: str = "Ativas") -> list[dict[str, Any]]:
        return nonconformities(self).rows(filter_text, state_filter)

    def quality_nc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return nonconformities(self).save(payload)

    def quality_nc_close(self, nc_id: str, eficacia: str = "") -> dict[str, Any]:
        return nonconformities(self).close(nc_id, eficacia)

    def quality_nc_release_material(self, nc_id: str, decision: str = "Aprovado pela qualidade") -> dict[str, Any]:
        return material_release(self).release(nc_id, decision)

    def quality_nc_remove(self, nc_id: str) -> None:
        nonconformities(self).remove(nc_id)

    def quality_document_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        return quality_documents(self).rows(filter_text)

    def quality_document_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return quality_documents(self).save(payload)

    def quality_document_remove(self, doc_id: str) -> None:
        return quality_documents(self).remove(doc_id)
