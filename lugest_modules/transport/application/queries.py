"""Transport projections computed from detached records."""
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class TransportReadRepository(Protocol):
    def trips(self) -> list[dict[str, Any]]: ...
    def orders(self) -> list[dict[str, Any]]: ...
    def quotes(self) -> list[dict[str, Any]]: ...
    def clients(self) -> list[dict[str, Any]]: ...
    def tariffs(self) -> list[dict[str, Any]]: ...
    def guides(self) -> list[dict[str, Any]]: ...
    def suppliers(self) -> list[dict[str, Any]]: ...
    def trip(self, number: str) -> dict[str, Any] | None: ...
    def order(self, number: str) -> dict[str, Any] | None: ...
    def quote(self, number: str) -> dict[str, Any] | None: ...
    def client(self, code: str) -> dict[str, Any]: ...

@dataclass(frozen=True)
class TransportQueryRules:
    parse_float: Callable
    norm_text: Callable
    normalize_supplier: Callable
    tariff_suggestion: Callable
    update_dispatch_state: Callable
    order_pieces: Callable
    available_dispatch_quantity: Callable
    checklist: Callable
    defaults: Callable

class TransportQueries:
    def __init__(self, repository: TransportReadRepository, rules: TransportQueryRules):
        self.repository = repository
        self.rules = rules

    def note(self, enc: dict[str, Any]) -> str:
        note = str((enc or {}).get("nota_transporte", "") or "").strip()
        if note:
            return note
        quote_num = str((enc or {}).get("numero_orcamento", "") or "").strip()
        if not quote_num:
            return ""
        quote = self.repository.quote(quote_num)
        if not isinstance(quote, dict):
            return ""
        return str(quote.get("nota_transporte", "") or "").strip()


    def mode(self, enc: dict[str, Any]) -> str:
        note = str(self.note(enc) or "").strip()
        note_norm = self.rules.norm_text(note)
        if "subcontrat" in note_norm:
            return "Subcontratado"
        if "nosso cargo" in note_norm or ("transporte" in note_norm and "nosso" in note_norm):
            return "Transporte a Nosso Cargo"
        if "cliente" in note_norm or "vosso cargo" in note_norm:
            return "Transporte a Cargo do Cliente"
        return note


    def zone(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> str:
        enc = dict(enc or {})
        zone = str(enc.get("zona_transporte", "") or "").strip()
        if zone:
            return zone
        cliente_obj = dict(cliente_obj or {})
        if not cliente_obj:
            cli_code = str(enc.get("cliente", "") or "").strip()
            if cli_code:
                cliente_obj = self.repository.client(cli_code) or {}
        for value in (
            cliente_obj.get("localidade", ""),
            cliente_obj.get("codigo_postal", ""),
        ):
            txt = str(value or "").strip()
            if txt:
                return txt
        return ""


    def zones(self) -> list[str]:
        values: list[str] = []
        for row in self.repository.tariffs():
            txt = str((row or {}).get("zona", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        for row in self.repository.orders():
            txt = self.zone(row)
            if txt and txt not in values:
                values.append(txt)
        for row in self.repository.quotes():
            txt = str((row or {}).get("zona_transporte", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        for row in self.repository.clients():
            txt = str((row or {}).get("localidade", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        values.sort(key=lambda item: self.rules.norm_text(item))
        return values


    def metrics(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> dict[str, Any]:
        enc = dict(enc or {})
        supplier_id, supplier_text, supplier_contact = self.rules.normalize_supplier(
            enc.get("transportadora_id", ""),
            enc.get("transportadora_nome", ""),
        )
        return {
            "modo": self.mode(enc),
            "paletes": round(self.rules.parse_float(enc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self.rules.parse_float(enc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self.rules.parse_float(enc.get("volume_m3", 0), 0), 3),
            "preco_transporte": round(self.rules.parse_float(enc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self.rules.parse_float(enc.get("custo_transporte", 0), 0), 2),
            "transportadora_id": supplier_id,
            "transportadora_nome": supplier_text,
            "transportadora_contacto": supplier_contact,
            "referencia_transporte": str(enc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": self.zone(enc, cliente_obj),
        }


    def summary(self, stops: list[dict[str, Any]], trip: dict[str, Any] | None = None) -> dict[str, float]:
        paletes_calc = round(sum(self.rules.parse_float(row.get("paletes", 0), 0) for row in list(stops or [])), 2)
        peso_calc = round(sum(self.rules.parse_float(row.get("peso_bruto_kg", 0), 0) for row in list(stops or [])), 2)
        volume_calc = round(sum(self.rules.parse_float(row.get("volume_m3", 0), 0) for row in list(stops or [])), 3)
        preco_total = round(sum(self.rules.parse_float(row.get("preco_transporte", 0), 0) for row in list(stops or [])), 2)
        custo_total = round(sum(self.rules.parse_float(row.get("custo_transporte", 0), 0) for row in list(stops or [])), 2)
        paletes = paletes_calc
        peso = peso_calc
        volume = volume_calc
        carga_manual = False
        if isinstance(trip, dict):
            custo_previsto = round(self.rules.parse_float(trip.get("custo_previsto", 0), 0), 2)
            if custo_previsto > 0:
                custo_total = custo_previsto
            paletes_manual = round(self.rules.parse_float(trip.get("paletes_total_manual", 0), 0), 2)
            peso_manual = round(self.rules.parse_float(trip.get("peso_total_manual_kg", 0), 0), 2)
            volume_manual = round(self.rules.parse_float(trip.get("volume_total_manual_m3", 0), 0), 3)
            if paletes_manual > 0:
                paletes = paletes_manual
                carga_manual = True
            if peso_manual > 0:
                peso = peso_manual
                carga_manual = True
            if volume_manual > 0:
                volume = volume_manual
                carga_manual = True
        return {
            "paletes": paletes,
            "peso_bruto_kg": peso,
            "volume_m3": volume,
            "paletes_calculadas": paletes_calc,
            "peso_bruto_kg_calculado": peso_calc,
            "volume_m3_calculado": volume_calc,
            "carga_manual": carga_manual,
            "preco_total": preco_total,
            "custo_total": custo_total,
            "margem_prevista": round(preco_total - custo_total, 2),
        }


    def own_cargo(self, enc: dict[str, Any]) -> bool:
        note_norm = self.rules.norm_text(self.note(enc))
        return (
            "nosso cargo" in note_norm
            or ("transporte" in note_norm and "nosso" in note_norm)
            or "subcontrat" in note_norm
        )


    def vehicles(self) -> list[str]:
        options: list[str] = []
        for tr in self.repository.trips():
            for value in (tr.get("viatura"), tr.get("matricula")):
                txt = str(value or "").strip()
                if txt and txt not in options:
                    options.append(txt)
        return options


    def drivers(self) -> list[str]:
        options: list[str] = []
        for tr in self.repository.trips():
            for value in (tr.get("motorista"), tr.get("telefone_motorista")):
                txt = str(value or "").strip()
                if txt and txt not in options:
                    options.append(txt)
        return options


    def latest_guide(self, order_num: str) -> dict[str, Any] | None:
        target = str(order_num or "").strip()
        if not target:
            return None
        matches = [
            dict(ex)
            for ex in self.repository.guides()
            if str((ex or {}).get("encomenda", "") or "").strip() == target and not bool((ex or {}).get("anulada"))
        ]
        if not matches:
            return None
        matches.sort(
            key=lambda row: (
                str(row.get("data_transporte", "") or row.get("data_emissao", "") or ""),
                str(row.get("numero", "") or ""),
            ),
            reverse=True,
        )
        return matches[0]


    def guides(self, order_num: str) -> list[dict[str, str]]:
        target = str(order_num or "").strip()
        if not target:
            return []
        rows = [
            {
                "numero": str(ex.get("numero", "") or "").strip(),
                "data_emissao": str(ex.get("data_emissao", "") or "").strip(),
                "data_transporte": str(ex.get("data_transporte", "") or "").strip(),
                "estado": str(ex.get("estado", "") or "").strip(),
                "local_descarga": str(ex.get("local_descarga", "") or "").strip(),
                "label": " | ".join(
                    [
                        part
                        for part in [
                            str(ex.get("numero", "") or "").strip(),
                            str(ex.get("data_transporte", "") or ex.get("data_emissao", "") or "").strip(),
                            str(ex.get("estado", "") or "").strip(),
                        ]
                        if part
                    ]
                ),
            }
            for ex in self.repository.guides()
            if str((ex or {}).get("encomenda", "") or "").strip() == target and not bool((ex or {}).get("anulada"))
        ]
        rows.sort(key=lambda row: (row.get("data_transporte") or row.get("data_emissao") or "", row.get("numero") or ""), reverse=True)
        return rows


    def stop_state(self, stop: dict[str, Any], trip_state: str = "") -> str:
        state = str((stop or {}).get("estado", "") or "").strip()
        if state:
            return state
        trip_txt = str(trip_state or "").strip()
        return trip_txt or "Planeada"


    def defaults(self) -> dict[str, Any]:
        payload = dict(self.rules.defaults())
        payload["vehicle_options"] = self.vehicles()
        payload["driver_options"] = self.drivers()
        payload["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.repository.suppliers() or [])]
        payload["zone_options"] = self.zones()
        return payload


    def pending_orders(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip(): str(tr.get("numero", "") or "").strip()
            for tr in self.repository.trips()
            if isinstance(tr, dict) and "anulad" not in self.rules.norm_text(str(tr.get("estado", "") or ""))
            for stop in list(tr.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        rows: list[dict[str, Any]] = []
        for enc in self.repository.orders():
            if not isinstance(enc, dict):
                continue
            enc_num = str(enc.get("numero", "") or "").strip()
            if not enc_num or not self.own_cargo(enc):
                continue
            self.rules.update_dispatch_state(enc)
            pieces = list(self.rules.order_pieces(enc))
            disponivel = sum(max(0.0, self.rules.parse_float(self.rules.available_dispatch_quantity(piece), 0)) for piece in pieces)
            latest_guide = self.latest_guide(enc_num) or {}
            if disponivel <= 0 and not latest_guide:
                continue
            if active_assignments.get(enc_num):
                continue
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = {}
            if cli_code:
                cli_obj = self.repository.client(cli_code) or {}
            cliente_txt = " - ".join([part for part in [cli_code, str(cli_obj.get("nome", "") or "").strip()] if part]).strip()
            metrics = self.metrics(enc, cli_obj)
            suggestion = self.rules.tariff_suggestion(
                metrics.get("transportadora_id", ""),
                metrics.get("transportadora_nome", ""),
                metrics.get("zona_transporte", ""),
                metrics.get("paletes", 0.0),
                metrics.get("peso_bruto_kg", 0.0),
                metrics.get("volume_m3", 0.0),
            )
            row = {
                "numero": enc_num,
                "cliente": cliente_txt or cli_code or "-",
                "cliente_codigo": cli_code,
                "estado": str(enc.get("estado", "") or "").strip(),
                "estado_expedicao": str(enc.get("estado_expedicao", "Nao expedida") or "Nao expedida").strip(),
                "estado_transporte": str(enc.get("estado_transporte", "") or "").strip(),
                "nota_transporte": metrics.get("modo", "") or self.note(enc),
                "preco_transporte": metrics.get("preco_transporte", 0.0),
                "custo_transporte": metrics.get("custo_transporte", 0.0),
                "paletes": metrics.get("paletes", 0.0),
                "peso_bruto_kg": metrics.get("peso_bruto_kg", 0.0),
                "volume_m3": metrics.get("volume_m3", 0.0),
                "transportadora_id": metrics.get("transportadora_id", ""),
                "transportadora_nome": metrics.get("transportadora_nome", ""),
                "referencia_transporte": metrics.get("referencia_transporte", ""),
                "zona_transporte": metrics.get("zona_transporte", ""),
                "local_descarga": str(enc.get("local_descarga", "") or cli_obj.get("morada", "") or "").strip(),
                "latitude": str(cli_obj.get("latitude", "") or "").strip(),
                "longitude": str(cli_obj.get("longitude", "") or "").strip(),
                "contacto": str(cli_obj.get("contacto", "") or "").strip(),
                "telefone": str(cli_obj.get("contacto", "") or "").strip(),
                "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                "guia_numero": str(latest_guide.get("numero", "") or "").strip(),
                "disponivel": round(disponivel, 1),
                "custo_sugerido": round(self.rules.parse_float(suggestion.get("custo_sugerido", 0), 0), 2),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("data_entrega") or "9999-99-99", item.get("numero") or ""))
        return rows


    def pending_overview(self) -> dict[str, int]:
        """Explain why orders do or do not appear in the transport planner."""

        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            for trip in self.repository.trips()
            if isinstance(trip, dict)
            and "anulad" not in self.rules.norm_text(str(trip.get("estado", "") or ""))
            for stop in list(trip.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        overview = {
            "total_orders": 0,
            "eligible": 0,
            "customer_transport": 0,
            "waiting_stock_or_guide": 0,
            "already_assigned": 0,
        }
        for order in self.repository.orders():
            if not isinstance(order, dict):
                continue
            order_number = str(order.get("numero", "") or "").strip()
            if not order_number:
                continue
            overview["total_orders"] += 1
            if not self.own_cargo(order):
                overview["customer_transport"] += 1
                continue
            if order_number in active_assignments:
                overview["already_assigned"] += 1
                continue
            pieces = list(self.rules.order_pieces(order))
            available = sum(
                max(0.0, self.rules.parse_float(self.rules.available_dispatch_quantity(piece), 0))
                for piece in pieces
            )
            if available <= 0 and not self.latest_guide(order_number):
                overview["waiting_stock_or_guide"] += 1
                continue
            overview["eligible"] += 1
        return overview


    def rows(self, filter_text: str = "", estado: str = "Todas") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_filter = str(estado or "Todas").strip().lower()
        rows: list[dict[str, Any]] = []
        for tr in self.repository.trips():
            if not isinstance(tr, dict):
                continue
            trip_state = str(tr.get("estado", "") or "Planeado").strip()
            if state_filter not in ("todas", "todos", "all", "") and trip_state.lower() != state_filter:
                continue
            stops = list(tr.get("paragens", []) or [])
            delivered = sum(1 for stop in stops if "entreg" in self.rules.norm_text(self.stop_state(stop, trip_state)))
            summary = self.summary(stops, tr)
            row = {
                "numero": str(tr.get("numero", "") or "").strip(),
                "data_planeada": str(tr.get("data_planeada", "") or "").strip(),
                "hora_saida": str(tr.get("hora_saida", "") or "").strip(),
                "tipo_responsavel": str(tr.get("tipo_responsavel", "") or "Nosso Cargo").strip(),
                "estado": trip_state,
                "pedido_transporte_estado": str(tr.get("pedido_transporte_estado", "") or "Nao pedido").strip() or "Nao pedido",
                "transportadora_nome": str(tr.get("transportadora_nome", "") or "").strip(),
                "viatura": str(tr.get("viatura", "") or tr.get("matricula", "") or "").strip(),
                "motorista": str(tr.get("motorista", "") or "").strip(),
                "matricula": str(tr.get("matricula", "") or "").strip(),
                "paragens": len(stops),
                "entregues": delivered,
                "pendentes": max(0, len(stops) - delivered),
                "paletes": summary.get("paletes", 0.0),
                "peso_bruto_kg": summary.get("peso_bruto_kg", 0.0),
                "preco_total": summary.get("preco_total", 0.0),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("data_planeada") or "", item.get("hora_saida") or "", item.get("numero") or ""), reverse=True)
        return rows


    def detail(self, numero: str) -> dict[str, Any]:
        trip = self.repository.trip(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        detail = {
            "numero": str(trip.get("numero", "") or "").strip(),
            "tipo_responsavel": str(trip.get("tipo_responsavel", "") or "Nosso Cargo").strip() or "Nosso Cargo",
            "estado": str(trip.get("estado", "") or "Planeado").strip() or "Planeado",
            "data_planeada": str(trip.get("data_planeada", "") or "").strip(),
            "hora_saida": str(trip.get("hora_saida", "") or "").strip(),
            "viatura": str(trip.get("viatura", "") or "").strip(),
            "matricula": str(trip.get("matricula", "") or "").strip(),
            "motorista": str(trip.get("motorista", "") or "").strip(),
            "telefone_motorista": str(trip.get("telefone_motorista", "") or "").strip(),
            "origem": str(trip.get("origem", "") or "").strip(),
            "origem_latitude": str(trip.get("origem_latitude", "") or "").strip(),
            "origem_longitude": str(trip.get("origem_longitude", "") or "").strip(),
            "transportadora_id": str(trip.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(trip.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(trip.get("referencia_transporte", "") or "").strip(),
            "custo_previsto": round(self.rules.parse_float(trip.get("custo_previsto", 0), 0), 2),
            "paletes_total_manual": round(self.rules.parse_float(trip.get("paletes_total_manual", 0), 0), 2),
            "peso_total_manual_kg": round(self.rules.parse_float(trip.get("peso_total_manual_kg", 0), 0), 2),
            "volume_total_manual_m3": round(self.rules.parse_float(trip.get("volume_total_manual_m3", 0), 0), 3),
            "pedido_transporte_estado": str(trip.get("pedido_transporte_estado", "") or "").strip() or "Nao pedido",
            "pedido_transporte_ref": str(trip.get("pedido_transporte_ref", "") or "").strip(),
            "pedido_transporte_at": str(trip.get("pedido_transporte_at", "") or "").strip(),
            "pedido_transporte_by": str(trip.get("pedido_transporte_by", "") or "").strip(),
            "pedido_transporte_obs": str(trip.get("pedido_transporte_obs", "") or "").strip(),
            "pedido_resposta_obs": str(trip.get("pedido_resposta_obs", "") or "").strip(),
            "pedido_confirmado_at": str(trip.get("pedido_confirmado_at", "") or "").strip(),
            "pedido_confirmado_by": str(trip.get("pedido_confirmado_by", "") or "").strip(),
            "pedido_recusado_at": str(trip.get("pedido_recusado_at", "") or "").strip(),
            "pedido_recusado_by": str(trip.get("pedido_recusado_by", "") or "").strip(),
            "observacoes": str(trip.get("observacoes", "") or "").strip(),
            "created_by": str(trip.get("created_by", "") or "").strip(),
            "created_at": str(trip.get("created_at", "") or "").strip(),
            "updated_at": str(trip.get("updated_at", "") or "").strip(),
            "paragens": [],
        }
        for stop in sorted(list(trip.get("paragens", []) or []), key=lambda row: int(self.rules.parse_float((row or {}).get("ordem", 0), 0) or 0)):
            if not isinstance(stop, dict):
                continue
            enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            enc = self.repository.order(enc_num) if enc_num else None
            cli_code = str(stop.get("cliente_codigo", "") or (enc or {}).get("cliente", "") or "").strip()
            cli_obj = {}
            if cli_code:
                cli_obj = self.repository.client(cli_code) or {}
            latest_guide = self.latest_guide(enc_num) or {}
            metrics = self.metrics(enc or {}, cli_obj)
            supplier_id, supplier_text, supplier_contact = self.rules.normalize_supplier(
                stop.get("transportadora_id", "") or detail.get("transportadora_id", "") or metrics.get("transportadora_id", ""),
                stop.get("transportadora_nome", "") or detail.get("transportadora_nome", "") or metrics.get("transportadora_nome", ""),
            )
            zone_txt = str(stop.get("zona_transporte", "") or metrics.get("zona_transporte", "") or "").strip()
            paletes_value = round(self.rules.parse_float(stop.get("paletes", metrics.get("paletes", 0)), 0), 2)
            peso_value = round(self.rules.parse_float(stop.get("peso_bruto_kg", metrics.get("peso_bruto_kg", 0)), 0), 2)
            volume_value = round(self.rules.parse_float(stop.get("volume_m3", metrics.get("volume_m3", 0)), 0), 3)
            suggestion = self.rules.tariff_suggestion(
                supplier_id,
                supplier_text,
                zone_txt,
                paletes_value,
                peso_value,
                volume_value,
            )
            detail["paragens"].append(
                {
                    "ordem": int(self.rules.parse_float(stop.get("ordem", 0), 0) or 0),
                    "encomenda_numero": enc_num,
                    "cliente_codigo": cli_code,
                    "cliente_nome": str(stop.get("cliente_nome", "") or cli_obj.get("nome", "") or "").strip(),
                    "zona_transporte": zone_txt,
                    "local_descarga": str(stop.get("local_descarga", "") or (enc or {}).get("local_descarga", "") or cli_obj.get("morada", "") or "").strip(),
                    "latitude": str(stop.get("latitude", "") or cli_obj.get("latitude", "") or "").strip(),
                    "longitude": str(stop.get("longitude", "") or cli_obj.get("longitude", "") or "").strip(),
                    "contacto": str(stop.get("contacto", "") or cli_obj.get("contacto", "") or "").strip(),
                    "telefone": str(stop.get("telefone", "") or cli_obj.get("contacto", "") or "").strip(),
                    "data_planeada": str(stop.get("data_planeada", "") or "").replace("T", " ")[:19],
                    "paletes": paletes_value,
                    "peso_bruto_kg": peso_value,
                    "volume_m3": volume_value,
                    "preco_transporte": round(self.rules.parse_float(stop.get("preco_transporte", metrics.get("preco_transporte", 0)), 0), 2),
                    "custo_transporte": round(self.rules.parse_float(stop.get("custo_transporte", metrics.get("custo_transporte", 0)), 0), 2),
                    "custo_manual": round(self.rules.parse_float(stop.get("custo_transporte", metrics.get("custo_transporte", 0)), 0), 2),
                    "custo_sugerido": round(self.rules.parse_float(suggestion.get("custo_sugerido", 0), 0), 2),
                    "tarifario_id": suggestion.get("tarifario_id", ""),
                    "tarifario_label": str(suggestion.get("tarifario_label", "") or "").strip(),
                    "transportadora_id": supplier_id,
                    "transportadora_nome": supplier_text,
                    "transportadora_contacto": supplier_contact,
                    "referencia_transporte": str(stop.get("referencia_transporte", "") or detail.get("referencia_transporte", "") or metrics.get("referencia_transporte", "") or "").strip(),
                    "nota_transporte": metrics.get("modo", "") or self.note(enc or {}),
                    "estado": self.stop_state(stop, detail["estado"]),
                    "check_carga_ok": bool(stop.get("check_carga_ok")),
                    "check_docs_ok": bool(stop.get("check_docs_ok")),
                    "check_paletes_ok": bool(stop.get("check_paletes_ok")),
                    "checklist_estado": self.rules.checklist(stop),
                    "pod_estado": str(stop.get("pod_estado", "") or "").strip(),
                    "pod_recebido_nome": str(stop.get("pod_recebido_nome", "") or "").strip(),
                    "pod_recebido_at": str(stop.get("pod_recebido_at", "") or "").replace("T", " ")[:19],
                    "pod_obs": str(stop.get("pod_obs", "") or "").strip(),
                    "observacoes": str(stop.get("observacoes", "") or "").strip(),
                    "guia_numero": str(stop.get("expedicao_numero", "") or latest_guide.get("numero", "") or "").strip(),
                    "estado_expedicao": str((enc or {}).get("estado_expedicao", "") or "").strip(),
                }
            )
        detail.update(self.summary(list(detail.get("paragens", []) or []), detail))
        detail["custo_sugerido_total"] = round(
            sum(self.rules.parse_float(stop.get("custo_sugerido", 0), 0) for stop in list(detail.get("paragens", []) or [])),
            2,
        )
        detail["checklist_ok"] = sum(1 for stop in list(detail.get("paragens", []) or []) if str(stop.get("checklist_estado", "") or "") == "OK")
        detail["pod_recebidos"] = sum(1 for stop in list(detail.get("paragens", []) or []) if "recebid" in self.rules.norm_text(str(stop.get("pod_estado", "") or "")))
        zones = []
        for stop in list(detail.get("paragens", []) or []):
            zone_txt = str(stop.get("zona_transporte", "") or "").strip()
            if zone_txt and zone_txt not in zones:
                zones.append(zone_txt)
        detail["zonas"] = zones
        detail["vehicle_options"] = self.vehicles()
        detail["driver_options"] = self.drivers()
        detail["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.repository.suppliers() or [])]
        detail["zone_options"] = self.zones()
        return detail


