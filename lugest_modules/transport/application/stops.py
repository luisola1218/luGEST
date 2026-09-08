"""Transport stop editing and progression rules with detached trip aggregates."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class TripRepository(Protocol):
    def get(self, number: str) -> dict[str, Any] | None: ...
    def save(self, trip: dict[str, Any], *, expected: dict[str, Any], sync_links: bool = False) -> None: ...

@dataclass(frozen=True)
class StopRules:
    parse_float: Callable
    norm_text: Callable
    now_iso: Callable
    actor: Callable
    normalize_supplier: Callable
    guide_options: Callable

class TransportStops:
    def __init__(self, repository: TripRepository, rules: StopRules):
        self.repository = repository
        self.rules = rules

    def checklist(self, stop: dict[str, Any]) -> str:
        checks = [
            bool((stop or {}).get("check_carga_ok")),
            bool((stop or {}).get("check_docs_ok")),
            bool((stop or {}).get("check_paletes_ok")),
        ]
        if checks and all(checks):
            return "OK"
        if any(checks):
            return "Parcial"
        return "Pendente"


    def reindex(self, trip: dict[str, Any]) -> None:
        stops = list(trip.get("paragens", []) or [])
        stops.sort(
            key=lambda row: (
                int(self.rules.parse_float((row or {}).get("ordem", 0), 0) or 0),
                str((row or {}).get("data_planeada", "") or ""),
                str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or ""),
            )
        )
        for index, stop in enumerate(stops, start=1):
            stop["ordem"] = index
        trip["paragens"] = stops


    def request_service(self, numero: str, payload: dict[str, Any] | None = None) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        payload = dict(payload or {})
        supplier_id, supplier_text, _supplier_contact = self.rules.normalize_supplier(
            payload.get("transportadora_id", trip.get("transportadora_id", "")),
            payload.get("transportadora_nome", trip.get("transportadora_nome", "")),
        )
        if "transportadora_id" in payload or "transportadora_nome" in payload:
            trip["transportadora_id"] = supplier_id
            trip["transportadora_nome"] = supplier_text
        trip["paletes_total_manual"] = round(self.rules.parse_float(payload.get("paletes_total_manual", trip.get("paletes_total_manual", 0)), 0), 2)
        trip["peso_total_manual_kg"] = round(self.rules.parse_float(payload.get("peso_total_manual_kg", trip.get("peso_total_manual_kg", 0)), 0), 2)
        trip["volume_total_manual_m3"] = round(self.rules.parse_float(payload.get("volume_total_manual_m3", trip.get("volume_total_manual_m3", 0)), 0), 3)
        trip["custo_previsto"] = round(self.rules.parse_float(payload.get("custo_previsto", trip.get("custo_previsto", 0)), 0), 2)
        request_state = str(payload.get("pedido_transporte_estado", trip.get("pedido_transporte_estado", "Pedido enviado")) or "Pedido enviado").strip() or "Pedido enviado"
        trip["pedido_transporte_estado"] = request_state
        trip["pedido_transporte_ref"] = str(payload.get("pedido_transporte_ref", trip.get("pedido_transporte_ref", "")) or "").strip()
        trip["pedido_transporte_obs"] = str(payload.get("pedido_transporte_obs", trip.get("pedido_transporte_obs", "")) or "").strip()
        trip["pedido_resposta_obs"] = str(payload.get("pedido_resposta_obs", trip.get("pedido_resposta_obs", "")) or "").strip()
        normalized_state = self.rules.norm_text(request_state)
        if normalized_state in {"nao pedido", "nao-pedido"}:
            trip["pedido_transporte_at"] = ""
            trip["pedido_transporte_by"] = ""
            trip["pedido_confirmado_at"] = ""
            trip["pedido_confirmado_by"] = ""
            trip["pedido_recusado_at"] = ""
            trip["pedido_recusado_by"] = ""
        else:
            trip["pedido_transporte_at"] = self.rules.now_iso()
            trip["pedido_transporte_by"] = self.rules.actor()
            if "confirm" in normalized_state:
                trip["pedido_confirmado_at"] = self.rules.now_iso()
                trip["pedido_confirmado_by"] = self.rules.actor()
                trip["pedido_recusado_at"] = ""
                trip["pedido_recusado_by"] = ""
            elif "recus" in normalized_state:
                trip["pedido_recusado_at"] = self.rules.now_iso()
                trip["pedido_recusado_by"] = self.rules.actor()
                trip["pedido_confirmado_at"] = ""
                trip["pedido_confirmado_by"] = ""
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original, sync_links=False)
        return str(trip["numero"])


    def update(self, numero: str, encomenda_numero: str, payload: dict[str, Any] | None = None) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        enc_num = str(encomenda_numero or "").strip()
        target = next(
            (
                row
                for row in list(trip.get("paragens", []) or [])
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() == enc_num
            ),
            None,
        )
        if target is None:
            raise ValueError("Paragem nao encontrada.")
        payload = dict(payload or {})
        guide_number = str(payload.get("expedicao_numero", target.get("expedicao_numero", "")) or "").strip()
        if guide_number:
            valid_guides = {row.get("numero", "") for row in self.rules.guide_options(enc_num)}
            if guide_number not in valid_guides:
                raise ValueError("A guia escolhida nao pertence a esta encomenda.")
        target["expedicao_numero"] = guide_number
        target["zona_transporte"] = str(payload.get("zona_transporte", target.get("zona_transporte", "")) or "").strip()
        target["local_descarga"] = str(payload.get("local_descarga", target.get("local_descarga", "")) or "").strip()
        target["latitude"] = str(payload.get("latitude", target.get("latitude", "")) or "").strip()
        target["longitude"] = str(payload.get("longitude", target.get("longitude", "")) or "").strip()
        target["contacto"] = str(payload.get("contacto", target.get("contacto", "")) or "").strip()
        target["telefone"] = str(payload.get("telefone", target.get("telefone", "")) or "").strip()
        target["data_planeada"] = str(payload.get("data_planeada", target.get("data_planeada", "")) or "").strip()
        if "check_carga_ok" in payload:
            target["check_carga_ok"] = bool(payload.get("check_carga_ok"))
        if "check_docs_ok" in payload:
            target["check_docs_ok"] = bool(payload.get("check_docs_ok"))
        if "check_paletes_ok" in payload:
            target["check_paletes_ok"] = bool(payload.get("check_paletes_ok"))
        if "pod_estado" in payload:
            target["pod_estado"] = str(payload.get("pod_estado", "") or "").strip()
        if "pod_recebido_nome" in payload:
            target["pod_recebido_nome"] = str(payload.get("pod_recebido_nome", "") or "").strip()
        if "pod_recebido_at" in payload:
            target["pod_recebido_at"] = str(payload.get("pod_recebido_at", "") or "").strip()
        if "pod_obs" in payload:
            target["pod_obs"] = str(payload.get("pod_obs", "") or "").strip()
        if "observacoes" in payload:
            target["observacoes"] = str(payload.get("observacoes", "") or "").strip()
        if (
            "recebid" in self.rules.norm_text(str(target.get("pod_estado", "") or ""))
            and not str(target.get("pod_recebido_at", "") or "").strip()
        ):
            target["pod_recebido_at"] = self.rules.now_iso()
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original, sync_links=False)
        return str(trip["numero"])


    def remove(self, numero: str, encomenda_numero: str) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        enc_num = str(encomenda_numero or "").strip()
        before = len(list(trip.get("paragens", []) or []))
        trip["paragens"] = [
            row
            for row in list(trip.get("paragens", []) or [])
            if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() != enc_num
        ]
        if len(list(trip.get("paragens", []) or [])) == before:
            raise ValueError("Paragem nao encontrada.")
        self.reindex(trip)
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original, sync_links=True)
        return str(trip["numero"])


    def move(self, numero: str, encomenda_numero: str, direction: int) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        enc_num = str(encomenda_numero or "").strip()
        stops = list(trip.get("paragens", []) or [])
        index = next(
            (
                idx
                for idx, row in enumerate(stops)
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() == enc_num
            ),
            -1,
        )
        if index < 0:
            raise ValueError("Paragem nao encontrada.")
        target = index + (1 if int(direction or 0) > 0 else -1)
        if target < 0 or target >= len(stops):
            return str(trip["numero"])
        stops[index], stops[target] = stops[target], stops[index]
        trip["paragens"] = stops
        for position, stop in enumerate(stops, start=1):
            stop["ordem"] = position
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original, sync_links=False)
        return str(trip["numero"])


    def set_status(self, numero: str, estado: str) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        state_txt = str(estado or "").strip()
        if not state_txt:
            raise ValueError("Estado obrigatorio.")
        state_norm = self.rules.norm_text(state_txt)
        stops = [row for row in list(trip.get("paragens", []) or []) if isinstance(row, dict)]
        if any(token in state_norm for token in ("carga", "transito", "conclu")) and not stops:
            raise ValueError("Adiciona pelo menos uma encomenda à viagem antes de avançar o estado.")
        if "transito" in state_norm:
            incomplete = [
                str(stop.get("encomenda_numero", "") or "").strip()
                for stop in stops
                if self.checklist(stop) != "OK"
                or not str(stop.get("expedicao_numero", "") or "").strip()
            ]
            if incomplete:
                raise ValueError(
                    "Antes de iniciar o transporte confirma carga, documentos, paletes e guia em: "
                    + ", ".join(incomplete)
                )
        if "conclu" in state_norm:
            incomplete = [
                str(stop.get("encomenda_numero", "") or "").strip()
                for stop in stops
                if "entreg" not in self.rules.norm_text(str(stop.get("estado", "") or ""))
                or "recebid" not in self.rules.norm_text(str(stop.get("pod_estado", "") or ""))
            ]
            if incomplete:
                raise ValueError(
                    "Só podes concluir a viagem depois de marcar cada destino como Entregue e registar o POD: "
                    + ", ".join(incomplete)
                )
        trip["estado"] = state_txt
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original, sync_links=True)
        return str(trip["numero"])


    def set_stop_status(self, numero: str, encomenda_numero: str, estado: str, observacoes: str = "") -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        enc_num = str(encomenda_numero or "").strip()
        target = next(
            (
                row
                for row in list(trip.get("paragens", []) or [])
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() == enc_num
            ),
            None,
        )
        if target is None:
            raise ValueError("Paragem nao encontrada.")
        next_state = str(estado or "").strip() or "Planeada"
        if "entreg" in self.rules.norm_text(next_state):
            missing: list[str] = []
            if not str(target.get("expedicao_numero", "") or "").strip():
                missing.append("guia")
            if self.checklist(target) != "OK":
                missing.append("checklist")
            if "recebid" not in self.rules.norm_text(str(target.get("pod_estado", "") or "")):
                missing.append("POD")
            if missing:
                raise ValueError("Antes de concluir o destino preenche: " + ", ".join(missing) + ".")
        target["estado"] = next_state
        if str(observacoes or "").strip():
            target["observacoes"] = str(observacoes or "").strip()
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original, sync_links=True)
        return str(trip["numero"])


