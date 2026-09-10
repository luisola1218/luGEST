"""Release quarantined material and close its NC as one prepared operation."""
import copy
from dataclasses import dataclass
from typing import Callable, Protocol

@dataclass
class ReleaseCatalogs:
    materials: list
    nonconformities: list
    notes: list

class ReleaseRepository(Protocol):
    def load(self) -> ReleaseCatalogs: ...
    def replace(self, catalogs: ReleaseCatalogs, *, expected: ReleaseCatalogs, event: dict, stock_event: tuple | None) -> None: ...

@dataclass(frozen=True)
class ReleaseRules:
    parse_float: Callable
    now_iso: Callable
    actor: Callable
    quarantine: Callable
    sync_notes: Callable
    release_movements: Callable | None = None

class MaterialRelease:
    def __init__(self, repository: ReleaseRepository, rules: ReleaseRules):
        self.repository = repository
        self.rules = rules

    def release(self, nc_id: str, decision: str = "Aprovado pela qualidade") -> dict:
        expected = self.repository.load()
        data = copy.deepcopy(expected)
        nc_id_txt = str(nc_id or "").strip()
        target = next((row for row in data.nonconformities if isinstance(row, dict) and str(row.get("id", "") or "").strip() == nc_id_txt), None)
        if target is None:
            raise ValueError("Nao conformidade nao encontrada.")
        material_id = str(target.get("material_id", "") or "").strip()
        if not material_id and str(target.get("entidade_tipo", "") or "").strip() == "Material":
            material_id = str(target.get("entidade_id", "") or "").strip()
        if not material_id:
            raise ValueError("Esta NC nao esta ligada a um material.")
        material = next((row for row in data.materials if str(row.get("id", "")).strip() == material_id), None)
        if material is None:
            raise ValueError("Material ligado a NC nao encontrado.")
        before_material = copy.deepcopy(material)
        now = self.rules.now_iso()
        linked = self.rules.release_movements(data, material, now, decision) if self.rules.release_movements else False
        if not linked:
            self.rules.quarantine(material, kind="Material")
        pending_qty = self.rules.parse_float(material.get("quality_pending_qty", 0), 0)
        before_qty = self.rules.parse_float(material.get("quantidade", 0), 0)
        stock_event = None
        if pending_qty > 0:
            material["quantidade"] = before_qty + pending_qty
            material["quality_pending_qty"] = 0.0
            material["quality_approved_qty"] = self.rules.parse_float(material.get("quality_approved_qty", 0), 0) + pending_qty
            stock_event = ("ENTRADA_QUALIDADE", f"{material_id} qtd={pending_qty} NC={nc_id_txt}", self.rules.actor())
        material["quality_status"] = "APROVADO"
        material["inspection_status"] = "APROVADO"
        material["quality_blocked"] = False
        material["inspection_decision"] = str(decision or "Aprovado pela qualidade").strip()
        material["quality_nc_id"] = ""
        material["supplier_claim_id"] = ""
        material["quality_released_at"] = now
        material["quality_released_by"] = self.rules.actor()
        material["atualizado_em"] = now
        target["decisao"] = str(decision or "Aprovado pela qualidade").strip()
        target["acao"] = (str(target.get("acao", "") or "").strip() + f"\nLibertacao de material: {material['inspection_decision']}").strip()
        target["estado"] = "Fechada"
        target["closed_at"] = now
        target["closed_by"] = self.rules.actor()
        target["updated_at"] = now
        target["updated_by"] = self.rules.actor()
        event = dict(
            action="Material libertado pela qualidade",
            entity_type="Material",
            entity_id=material_id,
            summary=f"NC {nc_id_txt}: {material['inspection_decision']}",
            before=before_material,
            after=material,
        )
        self.rules.sync_notes(data.notes, data.materials)
        self.repository.replace(data, expected=expected, event=event, stock_event=stock_event)
        return {"material_id": material_id, "quality_status": "APROVADO", "nc_id": nc_id_txt}

