from __future__ import annotations
from lugest_modules.transport.application.order_links import synchronize_orders
from lugest_qt.services.transport_composition import tariff_service, transport_stops, trip_commands, trip_assignments, transport_queries, route_report

import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any



class TransportBridgeMixin:
    """Transport route and tariff operations for the Qt bridge."""

    def _transport_defaults(self) -> dict[str, Any]:
        emit_cfg = dict(self.desktop_main.get_guia_emitente_info() or {})
        rodape = list(self.desktop_main.get_empresa_rodape_lines() or [])
        origem = str(
            emit_cfg.get("local_carga", "")
            or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))
            or ""
        ).strip()
        now_dt = datetime.now()
        return {
            "numero": "",
            "tipo_responsavel": "Nosso Cargo",
            "estado": "Planeado",
            "data_planeada": now_dt.date().isoformat(),
            "hora_saida": "08:00",
            "viatura": "",
            "matricula": "",
            "motorista": "",
            "telefone_motorista": "",
            "origem": origem,
            "origem_latitude": "",
            "origem_longitude": "",
            "paletes_total_manual": 0.0,
            "peso_total_manual_kg": 0.0,
            "volume_total_manual_m3": 0.0,
            "pedido_transporte_estado": "Nao pedido",
            "pedido_transporte_ref": "",
            "pedido_transporte_at": "",
            "pedido_transporte_by": "",
            "pedido_transporte_obs": "",
            "pedido_resposta_obs": "",
            "pedido_confirmado_at": "",
            "pedido_confirmado_by": "",
            "pedido_recusado_at": "",
            "pedido_recusado_by": "",
            "observacoes": "",
        }

    def _transport_note_for_order(self, enc: dict[str, Any]) -> str:
        return transport_queries(self).note(enc)

    def _transport_mode_for_order(self, enc: dict[str, Any]) -> str:
        return transport_queries(self).mode(enc)

    def _transport_zone_for_order(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> str:
        return transport_queries(self).zone(enc, cliente_obj)

    def transport_zone_options(self) -> list[str]:
        return transport_queries(self).zones()

    def transport_tariff_defaults(self) -> dict[str, Any]:
        return tariff_service(self).defaults()

    def _transport_tariff_signature(self, row: dict[str, Any] | None) -> str:
        return tariff_service(self).signature(row)

    def _transport_tariff_next_id(self) -> int:
        return tariff_service(self).next_id()

    def _transport_tariff_match(self, transportadora_id: Any = "", transportadora_nome: Any = "", zona: Any = "") -> dict[str, Any] | None:
        return tariff_service(self).match(transportadora_id, transportadora_nome, zona)

    def _transport_tariff_cost_from_row(self, row: dict[str, Any] | None, paletes: Any = 0, peso_bruto_kg: Any = 0, volume_m3: Any = 0) -> float:
        return tariff_service(self).cost(row, paletes, peso_bruto_kg, volume_m3)

    def _transport_tariff_suggestion(
        self,
        transportadora_id: Any = "",
        transportadora_nome: Any = "",
        zona: Any = "",
        paletes: Any = 0,
        peso_bruto_kg: Any = 0,
        volume_m3: Any = 0,
    ) -> dict[str, Any]:
        return tariff_service(self).suggestion(transportadora_id, transportadora_nome, zona, paletes, peso_bruto_kg, volume_m3)

    def _transport_metrics_for_order(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> dict[str, Any]:
        return transport_queries(self).metrics(enc, cliente_obj)

    def _transport_stop_summary(self, stops: list[dict[str, Any]], trip: dict[str, Any] | None = None) -> dict[str, float]:
        return transport_queries(self).summary(stops, trip)

    def _transport_is_own_cargo(self, enc: dict[str, Any]) -> bool:
        return transport_queries(self).own_cargo(enc)

    def _transport_vehicle_options(self) -> list[str]:
        return transport_queries(self).vehicles()

    def _transport_driver_options(self) -> list[str]:
        return transport_queries(self).drivers()

    def _transport_latest_guide_for_order(self, order_num: str) -> dict[str, Any] | None:
        return transport_queries(self).latest_guide(order_num)

    def transport_guide_options(self, order_num: str) -> list[dict[str, str]]:
        return transport_queries(self).guides(order_num)

    def _transport_find(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("transportes", []) or [])
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _transport_stop_state(self, stop: dict[str, Any], trip_state: str = "") -> str:
        return transport_queries(self).stop_state(stop, trip_state)

    def _transport_stop_checklist_state(self, stop: dict[str, Any]) -> str:
        return transport_stops(self).checklist(stop)

    def _transport_reindex_stops(self, trip: dict[str, Any]) -> None:
        return transport_stops(self).reindex(trip)

    def _transport_sync_order_links(self) -> None:
        data = self.ensure_data()
        orders = list(data.get("encomendas", []) or [])
        updated = synchronize_orders(list(data.get("transportes", []) or []), orders, self.desktop_main.norm_text)
        for original, projected in zip(orders, updated):
            if isinstance(original, dict):
                original.update(projected)

    def transport_defaults(self) -> dict[str, Any]:
        return transport_queries(self).defaults()

    def transport_tariff_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        return tariff_service(self).rows(filter_text)

    def transport_tariff_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return tariff_service(self).save(payload)

    def transport_tariff_remove(self, tariff_id: Any) -> None:
        return tariff_service(self).remove(tariff_id)

    def transport_pending_orders(self, filter_text: str = "") -> list[dict[str, Any]]:
        return transport_queries(self).pending_orders(filter_text)

    def transport_pending_overview(self) -> dict[str, int]:
        return transport_queries(self).pending_overview()

    def transport_rows(self, filter_text: str = "", estado: str = "Todas") -> list[dict[str, Any]]:
        return transport_queries(self).rows(filter_text, estado)

    def transport_detail(self, numero: str) -> dict[str, Any]:
        return transport_queries(self).detail(numero)

    def transport_create_or_update(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.transport_detail(trip_commands(self).save(payload))

    def transport_request_service(self, numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).request_service(numero, payload))

    def transport_remove_trip(self, numero: str) -> None:
        return trip_commands(self).remove(numero)

    def transport_update_stop(self, numero: str, encomenda_numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).update(numero, encomenda_numero, payload))

    def transport_assign_orders(self, numero: str, order_numbers: list[str]) -> dict[str, Any]:
        return self.transport_detail(trip_assignments(self).assign(numero, order_numbers))

    def transport_apply_suggested_cost(self, numero: str) -> dict[str, Any]:
        return self.transport_detail(trip_assignments(self).apply_suggested_cost(numero))

    def transport_remove_stop(self, numero: str, encomenda_numero: str) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).remove(numero, encomenda_numero))

    def transport_move_stop(self, numero: str, encomenda_numero: str, direction: int) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).move(numero, encomenda_numero, direction))

    def transport_set_status(self, numero: str, estado: str) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).set_status(numero, estado))

    def transport_set_stop_status(self, numero: str, encomenda_numero: str, estado: str, observacoes: str = "") -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).set_stop_status(numero, encomenda_numero, estado, observacoes))

    def transport_route_sheet_render(self, numero: str, path: str | Path) -> Path:
        return route_report(self).render(self.transport_detail(numero), path)

    def _transport_route_sheet_render_legacy(self, numero: str, path: str | Path) -> Path:
        return route_report(self).render_legacy(self.transport_detail(numero), path)

    def transport_route_sheet_open(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_transporte_{str(numero or '').strip()}.pdf"
        self.transport_route_sheet_render(numero, target)
        os.startfile(str(target))
        return target

