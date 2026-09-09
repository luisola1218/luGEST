"""Create and remove transport trips using prepared aggregates."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class TripCommandRepository(Protocol):
    def get(self, number: str) -> dict[str, Any] | None: ...
    def allocate_number(self, requested: str) -> str: ...
    def save(self, trip: dict[str, Any], *, expected: dict[str, Any] | None, sync_links: bool = True) -> None: ...
    def remove(self, number: str, *, expected: dict[str, Any]) -> None: ...

@dataclass(frozen=True)
class TripRules:
    parse_float: Callable
    norm_text: Callable
    now_iso: Callable
    actor: Callable
    normalize_supplier: Callable
    defaults: Callable
    reindex: Callable

class TripCommands:
    def __init__(self, repository: TripCommandRepository, rules: TripRules):
        self.repository = repository
        self.rules = rules

    def save(self, payload: dict[str, Any]) -> str:
        numero = str(payload.get("numero", "") or "").strip()
        trip = self.repository.get(numero) if numero else None
        original = deepcopy(trip)
        if trip is None:
            trip = {
                "numero": numero,
                "paragens": [],
                "created_by": self.rules.actor(),
                "created_at": self.rules.now_iso(),
            }
        defaults = self.rules.defaults()
        supplier_id, supplier_text, _supplier_contact = self.rules.normalize_supplier(
            payload.get("transportadora_id", trip.get("transportadora_id", "")),
            payload.get("transportadora_nome", trip.get("transportadora_nome", "")),
        )
        trip["tipo_responsavel"] = str(payload.get("tipo_responsavel", trip.get("tipo_responsavel", defaults["tipo_responsavel"])) or defaults["tipo_responsavel"]).strip() or defaults["tipo_responsavel"]
        trip["estado"] = str(payload.get("estado", trip.get("estado", defaults["estado"])) or defaults["estado"]).strip() or defaults["estado"]
        trip["data_planeada"] = str(payload.get("data_planeada", trip.get("data_planeada", defaults["data_planeada"])) or defaults["data_planeada"]).strip()
        trip["hora_saida"] = str(payload.get("hora_saida", trip.get("hora_saida", defaults["hora_saida"])) or defaults["hora_saida"]).strip()
        trip["viatura"] = str(payload.get("viatura", trip.get("viatura", "")) or "").strip()
        trip["matricula"] = str(payload.get("matricula", trip.get("matricula", "")) or "").strip()
        trip["motorista"] = str(payload.get("motorista", trip.get("motorista", "")) or "").strip()
        trip["telefone_motorista"] = str(payload.get("telefone_motorista", trip.get("telefone_motorista", "")) or "").strip()
        trip["origem"] = str(payload.get("origem", trip.get("origem", defaults["origem"])) or defaults["origem"]).strip()
        trip["origem_latitude"] = str(payload.get("origem_latitude", trip.get("origem_latitude", "")) or "").strip()
        trip["origem_longitude"] = str(payload.get("origem_longitude", trip.get("origem_longitude", "")) or "").strip()
        trip["transportadora_id"] = supplier_id
        trip["transportadora_nome"] = supplier_text
        trip["referencia_transporte"] = str(payload.get("referencia_transporte", trip.get("referencia_transporte", "")) or "").strip()
        trip["custo_previsto"] = round(self.rules.parse_float(payload.get("custo_previsto", trip.get("custo_previsto", 0)), 0), 2)
        trip["paletes_total_manual"] = round(self.rules.parse_float(payload.get("paletes_total_manual", trip.get("paletes_total_manual", 0)), 0), 2)
        trip["peso_total_manual_kg"] = round(self.rules.parse_float(payload.get("peso_total_manual_kg", trip.get("peso_total_manual_kg", 0)), 0), 2)
        trip["volume_total_manual_m3"] = round(self.rules.parse_float(payload.get("volume_total_manual_m3", trip.get("volume_total_manual_m3", 0)), 0), 3)
        trip["pedido_transporte_estado"] = str(payload.get("pedido_transporte_estado", trip.get("pedido_transporte_estado", "Nao pedido")) or "Nao pedido").strip() or "Nao pedido"
        trip["pedido_transporte_ref"] = str(payload.get("pedido_transporte_ref", trip.get("pedido_transporte_ref", "")) or "").strip()
        trip["pedido_transporte_at"] = str(payload.get("pedido_transporte_at", trip.get("pedido_transporte_at", "")) or "").strip()
        trip["pedido_transporte_by"] = str(payload.get("pedido_transporte_by", trip.get("pedido_transporte_by", "")) or "").strip()
        trip["pedido_transporte_obs"] = str(payload.get("pedido_transporte_obs", trip.get("pedido_transporte_obs", "")) or "").strip()
        trip["pedido_resposta_obs"] = str(payload.get("pedido_resposta_obs", trip.get("pedido_resposta_obs", "")) or "").strip()
        trip["pedido_confirmado_at"] = str(payload.get("pedido_confirmado_at", trip.get("pedido_confirmado_at", "")) or "").strip()
        trip["pedido_confirmado_by"] = str(payload.get("pedido_confirmado_by", trip.get("pedido_confirmado_by", "")) or "").strip()
        trip["pedido_recusado_at"] = str(payload.get("pedido_recusado_at", trip.get("pedido_recusado_at", "")) or "").strip()
        trip["pedido_recusado_by"] = str(payload.get("pedido_recusado_by", trip.get("pedido_recusado_by", "")) or "").strip()
        if "subcontrat" in self.rules.norm_text(trip["tipo_responsavel"]) and not trip["transportadora_nome"]:
            raise ValueError("Seleciona a transportadora externa para viagens subcontratadas.")
        trip["observacoes"] = str(payload.get("observacoes", trip.get("observacoes", "")) or "").strip()
        trip["updated_at"] = self.rules.now_iso()
        self.rules.reindex(trip)
        if original is None:
            numero = self.repository.allocate_number(numero)
            trip["numero"] = numero
        self.repository.save(trip, expected=original, sync_links=True)
        return numero


    def remove(self, numero: str) -> None:
        number = str(numero or "").strip()
        if not number:
            raise ValueError("Seleciona uma viagem.")
        trip = self.repository.get(number)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        self.repository.remove(number, expected=trip)
