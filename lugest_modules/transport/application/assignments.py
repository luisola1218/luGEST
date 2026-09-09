"""Prepare assignment batches and cost updates before publishing a trip."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol
from lugest_modules.transport.application.stops import TripRepository

class AssignmentRepository(TripRepository, Protocol):
    def all(self) -> list[dict[str, Any]]: ...
    def order(self, number: str) -> dict[str, Any] | None: ...
    def client(self, code: str) -> dict[str, Any]: ...
    def save_assignment(self, trip: dict[str, Any], *, expected: dict[str, Any], catalog: list[dict[str, Any]]) -> None: ...

@dataclass(frozen=True)
class AssignmentRules:
    parse_float: Callable
    norm_text: Callable
    now_iso: Callable
    own_cargo: Callable
    latest_guide: Callable
    metrics: Callable
    suggestion: Callable
    reindex: Callable
    detail: Callable

class TripAssignments:
    def __init__(self, repository: AssignmentRepository, rules: AssignmentRules):
        self.repository = repository
        self.rules = rules

    def assign(self, numero: str, order_numbers: list[str]) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        catalog = self.repository.all()
        trip_state = str(trip.get("estado", "") or "Planeado").strip()
        if "conclu" in self.rules.norm_text(trip_state) or "anulad" in self.rules.norm_text(trip_state):
            raise ValueError("Nao podes alterar uma viagem concluida ou anulada.")
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip(): str(tr.get("numero", "") or "").strip()
            for tr in catalog
            if isinstance(tr, dict) and "anulad" not in self.rules.norm_text(str(tr.get("estado", "") or ""))
            for stop in list(tr.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        existing_orders = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            for stop in list(trip.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        order_list = []
        for raw in list(order_numbers or []):
            num = str(raw or "").strip()
            if num and num not in order_list:
                order_list.append(num)
        if not order_list:
            raise ValueError("Seleciona pelo menos uma encomenda.")
        stop_dt = ""
        if str(trip.get("data_planeada", "") or "").strip():
            stop_dt = str(trip.get("data_planeada", "")).strip()
            if str(trip.get("hora_saida", "") or "").strip():
                stop_dt = f"{stop_dt}T{str(trip.get('hora_saida', '')).strip()}:00"
        for enc_num in order_list:
            if enc_num in existing_orders:
                continue
            assigned_trip = active_assignments.get(enc_num)
            if assigned_trip and assigned_trip != trip.get("numero"):
                raise ValueError(f"A encomenda {enc_num} ja esta afeta ao transporte {assigned_trip}.")
            enc = self.repository.order(enc_num)
            if enc is None:
                raise ValueError(f"Encomenda nao encontrada: {enc_num}")
            if not self.rules.own_cargo(enc):
                raise ValueError(f"A encomenda {enc_num} nao esta definida como transporte a nosso cargo.")
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = self.repository.client(cli_code) if cli_code else {}
            latest_guide = self.rules.latest_guide(enc_num) or {}
            metrics = self.rules.metrics(enc, cli_obj)
            carrier_id = str(trip.get("transportadora_id", "") or metrics.get("transportadora_id", "") or "").strip()
            carrier_name = str(trip.get("transportadora_nome", "") or metrics.get("transportadora_nome", "") or "").strip()
            zone_txt = str(metrics.get("zona_transporte", "") or "").strip()
            suggestion = self.rules.suggestion(
                carrier_id,
                carrier_name,
                zone_txt,
                metrics.get("paletes", 0.0),
                metrics.get("peso_bruto_kg", 0.0),
                metrics.get("volume_m3", 0.0),
            )
            order_cost = round(self.rules.parse_float(metrics.get("custo_transporte", 0), 0), 2)
            suggested_cost = round(self.rules.parse_float(suggestion.get("custo_sugerido", 0), 0), 2)
            trip.setdefault("paragens", []).append(
                {
                    "ordem": len(list(trip.get("paragens", []) or [])) + 1,
                    "encomenda_numero": enc_num,
                    "expedicao_numero": str(latest_guide.get("numero", "") or "").strip(),
                    "cliente_codigo": cli_code,
                    "cliente_nome": str(cli_obj.get("nome", "") or "").strip(),
                    "zona_transporte": zone_txt,
                    "local_descarga": str(enc.get("local_descarga", "") or cli_obj.get("morada", "") or "").strip(),
                    "latitude": str(cli_obj.get("latitude", "") or "").strip(),
                    "longitude": str(cli_obj.get("longitude", "") or "").strip(),
                    "contacto": str(cli_obj.get("contacto", "") or "").strip(),
                    "telefone": str(cli_obj.get("contacto", "") or "").strip(),
                    "data_planeada": stop_dt,
                    "paletes": metrics.get("paletes", 0.0),
                    "peso_bruto_kg": metrics.get("peso_bruto_kg", 0.0),
                    "volume_m3": metrics.get("volume_m3", 0.0),
                    "preco_transporte": metrics.get("preco_transporte", 0.0),
                    "custo_transporte": order_cost if order_cost > 0 else suggested_cost,
                    "transportadora_id": carrier_id,
                    "transportadora_nome": carrier_name,
                    "referencia_transporte": str(metrics.get("referencia_transporte", "") or trip.get("referencia_transporte", "") or "").strip(),
                    "check_carga_ok": False,
                    "check_docs_ok": False,
                    "check_paletes_ok": False,
                    "pod_estado": "",
                    "pod_recebido_nome": "",
                    "pod_recebido_at": "",
                    "pod_obs": "",
                    "estado": "Planeada",
                    "observacoes": "",
                }
            )
        self.rules.reindex(trip)
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save_assignment(trip, expected=original, catalog=catalog)
        return str(trip["numero"])


    def apply_suggested_cost(self, numero: str) -> str:
        trip = self.repository.get(str(numero or "").strip())
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        original = deepcopy(trip)
        detail = self.rules.detail(numero)
        suggested_map = {
            str(stop.get("encomenda_numero", "") or "").strip(): round(self.rules.parse_float(stop.get("custo_sugerido", 0), 0), 2)
            for stop in list(detail.get("paragens", []) or [])
        }
        total = 0.0
        applied = 0
        for stop in list(trip.get("paragens", []) or []):
            if not isinstance(stop, dict):
                continue
            enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            suggested = round(self.rules.parse_float(suggested_map.get(enc_num, 0), 0), 2)
            if suggested <= 0:
                continue
            stop["custo_transporte"] = suggested
            total += suggested
            applied += 1
        if applied <= 0:
            raise ValueError("Sem custos sugeridos para aplicar nesta viagem.")
        trip["custo_previsto"] = round(total, 2)
        trip["updated_at"] = self.rules.now_iso()
        self.repository.save(trip, expected=original)
        return str(trip["numero"])


