"""Read delivery decisions and reconcile quality totals on owned catalogs."""
from typing import Any
from .stock_policy import status_code

class DeliveryMovements:
    def __init__(self, parse_float, is_material, nonconformities):
        self.parse_float = parse_float
        self.is_material = is_material
        self.nonconformities = nonconformities

    def pending_quantity(self, movement: dict[str, Any]) -> float:
        qty = self.parse_float(movement.get("qtd", 0), 0)
        approved = self.parse_float(movement.get("quality_approved_qty", 0), 0)
        rejected = self.parse_float(movement.get("quality_rejected_qty", 0), 0)
        status = status_code(movement.get("quality_status", movement.get("inspection_status", "")))
        if approved <= 0 and rejected <= 0:
            if status == "APROVADO":
                approved = qty
            elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                rejected = qty
        return max(0.0, qty - approved - rejected)

    def movement_id(self, note_number: str, line_index: int, movement_index: int, entity_type: str, entity_id: str) -> str:
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

    def rows(self, notes) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        counts = {}
        for note in notes:
            for line in (note.get("linhas", []) or []):
                kind = "Material" if self.is_material(line.get("origem", "")) else "Produto"
                for movement in (line.get("entregas_linha", []) or []):
                    key = (str(note.get("numero", "")).strip(), kind, str(movement.get("stock_ref") or line.get("ref", "")).strip())
                    counts[key] = counts.get(key, 0) + 1
        for note in notes:
            if not isinstance(note, dict):
                continue
            note_number = str(note.get("numero", "") or "").strip()
            for line_index, line in enumerate(list(note.get("linhas", []) or [])):
                if not isinstance(line, dict):
                    continue
                entity_type = "Material" if self.is_material(line.get("origem", "")) else "Produto"
                fallback_ref = str(line.get("ref", "") or "").strip()
                for movement_index, movement in enumerate(list(line.get("entregas_linha", []) or [])):
                    if not isinstance(movement, dict):
                        continue
                    entity_id = str(movement.get("stock_ref", "") or fallback_ref).strip()
                    if not entity_id:
                        continue
                    movement_id = str(movement.get("quality_movement_id", "") or "").strip()
                    if not movement_id:
                        movement_id = self.movement_id(note_number, line_index, movement_index, entity_type, entity_id)
                    status = status_code(
                        movement.get("quality_status", "")
                        or movement.get("inspection_status", "")
                        or line.get("quality_status", "")
                        or line.get("inspection_status", "")
                        or "EM_INSPECAO"
                    )
                    qty = self.parse_float(movement.get("qtd", 0), 0)
                    approved = self.parse_float(movement.get("quality_approved_qty", 0), 0)
                    rejected = self.parse_float(movement.get("quality_rejected_qty", 0), 0)
                    if approved <= 0 and rejected <= 0:
                        if status == "APROVADO":
                            approved = qty
                        elif status in {"REJEITADO", "DEVOLVER_FORNECEDOR"}:
                            rejected = qty
                        else:
                            open_nc = self.nonconformities.find_open(
                                {
                                    "origem": "Receção fornecedor",
                                    "referencia": note_number,
                                    "entidade_tipo": entity_type,
                                    "entidade_id": entity_id,
                                }
                            )
                            explicit = "quality_approved_qty" in movement or "quality_rejected_qty" in movement
                            linked = str((open_nc or {}).get("movement_id", "") or "").strip()
                            unambiguous = linked == movement_id if linked else counts.get((note_number, entity_type, entity_id), 0) == 1
                            if open_nc is not None and not explicit and unambiguous:
                                approved = min(qty, self.nonconformities.quantity(open_nc, "qtd_aprovada"))
                                rejected = min(qty - approved, self.nonconformities.quantity(open_nc, "qtd_rejeitada"))
                                if approved > 0 or rejected > 0:
                                    pending_guess = max(0.0, qty - approved - rejected)
                                    status = "EM_INSPECAO" if pending_guess > 0 else ("APROVADO" if approved > 0 else "REJEITADO")
                    pending = max(0.0, qty - approved - rejected)
                    rows.append(
                        {
                            "movement_id": movement_id,
                            "note": note,
                            "line": line,
                            "movement": movement,
                            "note_number": note_number,
                            "line_index": line_index,
                            "movement_index": movement_index,
                            "entity_type": entity_type,
                            "entity_id": entity_id,
                            "qty": qty,
                            "approved_qty": approved,
                            "rejected_qty": rejected,
                            "pending_qty": pending,
                            "status": status,
                        }
                    )
        return rows

    def reconcile(self, notes, materials, products) -> bool:
        movement_rows = self.rows(notes)
        totals: dict[tuple[str, str], dict[str, float]] = {}
        for row in movement_rows:
            key = (str(row.get("entity_type", "") or ""), str(row.get("entity_id", "") or ""))
            bucket = totals.setdefault(key, {"received": 0.0, "pending": 0.0, "approved": 0.0, "rejected": 0.0})
            bucket["received"] += self.parse_float(row.get("qty", 0), 0)
            bucket["pending"] += self.parse_float(row.get("pending_qty", 0), 0)
            bucket["approved"] += self.parse_float(row.get("approved_qty", 0), 0)
            bucket["rejected"] += self.parse_float(row.get("rejected_qty", 0), 0)
        changed = False

        def _apply(item: dict[str, Any], entity_type: str, entity_id: str) -> None:
            nonlocal changed
            total = totals.get((entity_type, entity_id))
            if not total:
                return
            pending = round(total["pending"], 4)
            received = round(total["received"], 4)
            approved = round(total["approved"], 4)
            rejected = round(total["rejected"], 4)
            updates = {
                "quality_pending_qty": pending,
                "quality_received_qty": received,
                "quality_approved_qty": approved,
                "quality_rejected_qty": rejected,
            }
            for key, value in updates.items():
                if abs(self.parse_float(item.get(key, 0), 0) - value) > 1e-6:
                    item[key] = value
                    changed = True
            current_status = status_code(item.get("quality_status", item.get("inspection_status", "")))
            if pending > 0:
                target_status = "EM_AVERIGUACAO" if current_status == "EM_AVERIGUACAO" else "EM_INSPECAO"
                target_blocked = True
            elif approved > 0:
                target_status = "APROVADO"
                target_blocked = False
            elif rejected > 0:
                target_status = "REJEITADO"
                target_blocked = True
            else:
                target_status = current_status
                target_blocked = current_status != "APROVADO"
            if str(item.get("quality_status", "") or "").strip() != target_status:
                item["quality_status"] = target_status
                item["inspection_status"] = target_status
                changed = True
            if bool(item.get("quality_blocked")) != target_blocked:
                item["quality_blocked"] = target_blocked
                changed = True

        for material in materials:
            if isinstance(material, dict):
                _apply(material, "Material", str(material.get("id", "") or "").strip())
        for product in products:
            if isinstance(product, dict):
                _apply(product, "Produto", str(product.get("codigo", "") or "").strip())
        return changed


    def release_material(self, notes, material, *, now, actor, decision):
        """Resolve pending receipts together with a material release; leave rejected units rejected."""
        identifier = str(material.get("id", "")).strip()
        rows = [row for row in self.rows(notes) if row["entity_type"] == "Material" and row["entity_id"] == identifier]
        if not rows:
            return False
        self.reconcile(notes, [material], [])
        for row in rows:
            movement = row["movement"]
            pending = row["pending_qty"]
            if pending <= 0:
                continue
            movement.update(quality_movement_id=row["movement_id"], stock_ref=identifier,
                            quality_approved_qty=row["approved_qty"] + pending,
                            quality_rejected_qty=row["rejected_qty"], quality_pending_qty=0,
                            quality_status="APROVADO", inspection_status="APROVADO",
                            inspection_at=now, inspection_by=actor, inspection_decision=decision,
                            quality_nc_id="")
            line = row["line"]
            line.update(quality_status="APROVADO", inspection_status="APROVADO", _stock_in=True)
        return True
