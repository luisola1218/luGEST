from __future__ import annotations
from lugest_modules.transport.application.order_links import synchronize_orders
from lugest_qt.services.transport_composition import tariff_service, transport_stops

import os
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text


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
        note = str((enc or {}).get("nota_transporte", "") or "").strip()
        if note:
            return note
        quote_num = str((enc or {}).get("numero_orcamento", "") or "").strip()
        if not quote_num:
            return ""
        quote = self._billing_quote_by_number(quote_num)
        if not isinstance(quote, dict):
            return ""
        return str(quote.get("nota_transporte", "") or "").strip()

    def _transport_mode_for_order(self, enc: dict[str, Any]) -> str:
        note = str(self._transport_note_for_order(enc) or "").strip()
        note_norm = self.desktop_main.norm_text(note)
        if "subcontrat" in note_norm:
            return "Subcontratado"
        if "nosso cargo" in note_norm or ("transporte" in note_norm and "nosso" in note_norm):
            return "Transporte a Nosso Cargo"
        if "cliente" in note_norm or "vosso cargo" in note_norm:
            return "Transporte a Cargo do Cliente"
        return note

    def _transport_zone_for_order(self, enc: dict[str, Any] | None, cliente_obj: dict[str, Any] | None = None) -> str:
        enc = dict(enc or {})
        zone = str(enc.get("zona_transporte", "") or "").strip()
        if zone:
            return zone
        cliente_obj = dict(cliente_obj or {})
        if not cliente_obj:
            cli_code = str(enc.get("cliente", "") or "").strip()
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn) and cli_code:
                cliente_obj = find_cliente_fn(self.ensure_data(), cli_code) or {}
        for value in (
            cliente_obj.get("localidade", ""),
            cliente_obj.get("codigo_postal", ""),
        ):
            txt = str(value or "").strip()
            if txt:
                return txt
        return ""

    def transport_zone_options(self) -> list[str]:
        values: list[str] = []
        for row in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            txt = str((row or {}).get("zona", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        for row in list(self.ensure_data().get("encomendas", []) or []):
            txt = self._transport_zone_for_order(row)
            if txt and txt not in values:
                values.append(txt)
        for row in list(self.ensure_data().get("orcamentos", []) or []):
            txt = str((row or {}).get("zona_transporte", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        for row in list(self.ensure_data().get("clientes", []) or []):
            txt = str((row or {}).get("localidade", "") or "").strip()
            if txt and txt not in values:
                values.append(txt)
        values.sort(key=lambda item: self.desktop_main.norm_text(item))
        return values

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
        enc = dict(enc or {})
        supplier_id, supplier_text, supplier_contact = self._normalize_supplier_reference(
            enc.get("transportadora_id", ""),
            enc.get("transportadora_nome", ""),
        )
        return {
            "modo": self._transport_mode_for_order(enc),
            "paletes": round(self._parse_float(enc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(enc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(enc.get("volume_m3", 0), 0), 3),
            "preco_transporte": round(self._parse_float(enc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(enc.get("custo_transporte", 0), 0), 2),
            "transportadora_id": supplier_id,
            "transportadora_nome": supplier_text,
            "transportadora_contacto": supplier_contact,
            "referencia_transporte": str(enc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": self._transport_zone_for_order(enc, cliente_obj),
        }

    def _transport_stop_summary(self, stops: list[dict[str, Any]], trip: dict[str, Any] | None = None) -> dict[str, float]:
        paletes_calc = round(sum(self._parse_float(row.get("paletes", 0), 0) for row in list(stops or [])), 2)
        peso_calc = round(sum(self._parse_float(row.get("peso_bruto_kg", 0), 0) for row in list(stops or [])), 2)
        volume_calc = round(sum(self._parse_float(row.get("volume_m3", 0), 0) for row in list(stops or [])), 3)
        preco_total = round(sum(self._parse_float(row.get("preco_transporte", 0), 0) for row in list(stops or [])), 2)
        custo_total = round(sum(self._parse_float(row.get("custo_transporte", 0), 0) for row in list(stops or [])), 2)
        paletes = paletes_calc
        peso = peso_calc
        volume = volume_calc
        carga_manual = False
        if isinstance(trip, dict):
            custo_previsto = round(self._parse_float(trip.get("custo_previsto", 0), 0), 2)
            if custo_previsto > 0:
                custo_total = custo_previsto
            paletes_manual = round(self._parse_float(trip.get("paletes_total_manual", 0), 0), 2)
            peso_manual = round(self._parse_float(trip.get("peso_total_manual_kg", 0), 0), 2)
            volume_manual = round(self._parse_float(trip.get("volume_total_manual_m3", 0), 0), 3)
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

    def _transport_is_own_cargo(self, enc: dict[str, Any]) -> bool:
        note_norm = self.desktop_main.norm_text(self._transport_note_for_order(enc))
        return (
            "nosso cargo" in note_norm
            or ("transporte" in note_norm and "nosso" in note_norm)
            or "subcontrat" in note_norm
        )

    def _transport_vehicle_options(self) -> list[str]:
        options: list[str] = []
        for tr in list(self.ensure_data().get("transportes", []) or []):
            for value in (tr.get("viatura"), tr.get("matricula")):
                txt = str(value or "").strip()
                if txt and txt not in options:
                    options.append(txt)
        return options

    def _transport_driver_options(self) -> list[str]:
        options: list[str] = []
        for tr in list(self.ensure_data().get("transportes", []) or []):
            for value in (tr.get("motorista"), tr.get("telefone_motorista")):
                txt = str(value or "").strip()
                if txt and txt not in options:
                    options.append(txt)
        return options

    def _transport_latest_guide_for_order(self, order_num: str) -> dict[str, Any] | None:
        target = str(order_num or "").strip()
        if not target:
            return None
        matches = [
            dict(ex)
            for ex in list(self.ensure_data().get("expedicoes", []) or [])
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

    def transport_guide_options(self, order_num: str) -> list[dict[str, str]]:
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
            for ex in list(self.ensure_data().get("expedicoes", []) or [])
            if str((ex or {}).get("encomenda", "") or "").strip() == target and not bool((ex or {}).get("anulada"))
        ]
        rows.sort(key=lambda row: (row.get("data_transporte") or row.get("data_emissao") or "", row.get("numero") or ""), reverse=True)
        return rows

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
        state = str((stop or {}).get("estado", "") or "").strip()
        if state:
            return state
        trip_txt = str(trip_state or "").strip()
        return trip_txt or "Planeada"

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
        payload = dict(self._transport_defaults())
        payload["vehicle_options"] = self._transport_vehicle_options()
        payload["driver_options"] = self._transport_driver_options()
        payload["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.ne_suppliers() or [])]
        payload["zone_options"] = self.transport_zone_options()
        return payload

    def transport_tariff_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        return tariff_service(self).rows(filter_text)

    def transport_tariff_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return tariff_service(self).save(payload)

    def transport_tariff_remove(self, tariff_id: Any) -> None:
        return tariff_service(self).remove(tariff_id)

    def transport_pending_orders(self, filter_text: str = "") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip(): str(tr.get("numero", "") or "").strip()
            for tr in list(data.get("transportes", []) or [])
            if isinstance(tr, dict) and "anulad" not in self.desktop_main.norm_text(str(tr.get("estado", "") or ""))
            for stop in list(tr.get("paragens", []) or [])
            if isinstance(stop, dict)
        }
        rows: list[dict[str, Any]] = []
        for enc in list(data.get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            enc_num = str(enc.get("numero", "") or "").strip()
            if not enc_num or not self._transport_is_own_cargo(enc):
                continue
            self.desktop_main.update_estado_expedicao_encomenda(enc)
            pieces = list(self.desktop_main.encomenda_pecas(enc))
            disponivel = sum(max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(piece), 0)) for piece in pieces)
            latest_guide = self._transport_latest_guide_for_order(enc_num) or {}
            if disponivel <= 0 and not latest_guide:
                continue
            if active_assignments.get(enc_num):
                continue
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = {}
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn):
                cli_obj = find_cliente_fn(data, cli_code) or {}
            cliente_txt = " - ".join([part for part in [cli_code, str(cli_obj.get("nome", "") or "").strip()] if part]).strip()
            metrics = self._transport_metrics_for_order(enc, cli_obj)
            suggestion = self._transport_tariff_suggestion(
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
                "nota_transporte": metrics.get("modo", "") or self._transport_note_for_order(enc),
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
                "custo_sugerido": round(self._parse_float(suggestion.get("custo_sugerido", 0), 0), 2),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("data_entrega") or "9999-99-99", item.get("numero") or ""))
        return rows

    def transport_pending_overview(self) -> dict[str, int]:
        """Explain why orders do or do not appear in the transport planner."""

        data = self.ensure_data()
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            for trip in list(data.get("transportes", []) or [])
            if isinstance(trip, dict)
            and "anulad" not in self.desktop_main.norm_text(str(trip.get("estado", "") or ""))
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
        for order in list(data.get("encomendas", []) or []):
            if not isinstance(order, dict):
                continue
            order_number = str(order.get("numero", "") or "").strip()
            if not order_number:
                continue
            overview["total_orders"] += 1
            if not self._transport_is_own_cargo(order):
                overview["customer_transport"] += 1
                continue
            if order_number in active_assignments:
                overview["already_assigned"] += 1
                continue
            pieces = list(self.desktop_main.encomenda_pecas(order))
            available = sum(
                max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(piece), 0))
                for piece in pieces
            )
            if available <= 0 and not self._transport_latest_guide_for_order(order_number):
                overview["waiting_stock_or_guide"] += 1
                continue
            overview["eligible"] += 1
        return overview

    def transport_rows(self, filter_text: str = "", estado: str = "Todas") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_filter = str(estado or "Todas").strip().lower()
        rows: list[dict[str, Any]] = []
        for tr in list(self.ensure_data().get("transportes", []) or []):
            if not isinstance(tr, dict):
                continue
            trip_state = str(tr.get("estado", "") or "Planeado").strip()
            if state_filter not in ("todas", "todos", "all", "") and trip_state.lower() != state_filter:
                continue
            stops = list(tr.get("paragens", []) or [])
            delivered = sum(1 for stop in stops if "entreg" in self.desktop_main.norm_text(self._transport_stop_state(stop, trip_state)))
            summary = self._transport_stop_summary(stops, tr)
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

    def transport_detail(self, numero: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
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
            "custo_previsto": round(self._parse_float(trip.get("custo_previsto", 0), 0), 2),
            "paletes_total_manual": round(self._parse_float(trip.get("paletes_total_manual", 0), 0), 2),
            "peso_total_manual_kg": round(self._parse_float(trip.get("peso_total_manual_kg", 0), 0), 2),
            "volume_total_manual_m3": round(self._parse_float(trip.get("volume_total_manual_m3", 0), 0), 3),
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
        for stop in sorted(list(trip.get("paragens", []) or []), key=lambda row: int(self._parse_float((row or {}).get("ordem", 0), 0) or 0)):
            if not isinstance(stop, dict):
                continue
            enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            enc = self.get_encomenda_by_numero(enc_num) if enc_num else None
            cli_code = str(stop.get("cliente_codigo", "") or (enc or {}).get("cliente", "") or "").strip()
            cli_obj = {}
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn) and cli_code:
                cli_obj = find_cliente_fn(self.ensure_data(), cli_code) or {}
            latest_guide = self._transport_latest_guide_for_order(enc_num) or {}
            metrics = self._transport_metrics_for_order(enc or {}, cli_obj)
            supplier_id, supplier_text, supplier_contact = self._normalize_supplier_reference(
                stop.get("transportadora_id", "") or detail.get("transportadora_id", "") or metrics.get("transportadora_id", ""),
                stop.get("transportadora_nome", "") or detail.get("transportadora_nome", "") or metrics.get("transportadora_nome", ""),
            )
            zone_txt = str(stop.get("zona_transporte", "") or metrics.get("zona_transporte", "") or "").strip()
            paletes_value = round(self._parse_float(stop.get("paletes", metrics.get("paletes", 0)), 0), 2)
            peso_value = round(self._parse_float(stop.get("peso_bruto_kg", metrics.get("peso_bruto_kg", 0)), 0), 2)
            volume_value = round(self._parse_float(stop.get("volume_m3", metrics.get("volume_m3", 0)), 0), 3)
            suggestion = self._transport_tariff_suggestion(
                supplier_id,
                supplier_text,
                zone_txt,
                paletes_value,
                peso_value,
                volume_value,
            )
            detail["paragens"].append(
                {
                    "ordem": int(self._parse_float(stop.get("ordem", 0), 0) or 0),
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
                    "preco_transporte": round(self._parse_float(stop.get("preco_transporte", metrics.get("preco_transporte", 0)), 0), 2),
                    "custo_transporte": round(self._parse_float(stop.get("custo_transporte", metrics.get("custo_transporte", 0)), 0), 2),
                    "custo_manual": round(self._parse_float(stop.get("custo_transporte", metrics.get("custo_transporte", 0)), 0), 2),
                    "custo_sugerido": round(self._parse_float(suggestion.get("custo_sugerido", 0), 0), 2),
                    "tarifario_id": suggestion.get("tarifario_id", ""),
                    "tarifario_label": str(suggestion.get("tarifario_label", "") or "").strip(),
                    "transportadora_id": supplier_id,
                    "transportadora_nome": supplier_text,
                    "transportadora_contacto": supplier_contact,
                    "referencia_transporte": str(stop.get("referencia_transporte", "") or detail.get("referencia_transporte", "") or metrics.get("referencia_transporte", "") or "").strip(),
                    "nota_transporte": metrics.get("modo", "") or self._transport_note_for_order(enc or {}),
                    "estado": self._transport_stop_state(stop, detail["estado"]),
                    "check_carga_ok": bool(stop.get("check_carga_ok")),
                    "check_docs_ok": bool(stop.get("check_docs_ok")),
                    "check_paletes_ok": bool(stop.get("check_paletes_ok")),
                    "checklist_estado": self._transport_stop_checklist_state(stop),
                    "pod_estado": str(stop.get("pod_estado", "") or "").strip(),
                    "pod_recebido_nome": str(stop.get("pod_recebido_nome", "") or "").strip(),
                    "pod_recebido_at": str(stop.get("pod_recebido_at", "") or "").replace("T", " ")[:19],
                    "pod_obs": str(stop.get("pod_obs", "") or "").strip(),
                    "observacoes": str(stop.get("observacoes", "") or "").strip(),
                    "guia_numero": str(stop.get("expedicao_numero", "") or latest_guide.get("numero", "") or "").strip(),
                    "estado_expedicao": str((enc or {}).get("estado_expedicao", "") or "").strip(),
                }
            )
        detail.update(self._transport_stop_summary(list(detail.get("paragens", []) or []), detail))
        detail["custo_sugerido_total"] = round(
            sum(self._parse_float(stop.get("custo_sugerido", 0), 0) for stop in list(detail.get("paragens", []) or [])),
            2,
        )
        detail["checklist_ok"] = sum(1 for stop in list(detail.get("paragens", []) or []) if str(stop.get("checklist_estado", "") or "") == "OK")
        detail["pod_recebidos"] = sum(1 for stop in list(detail.get("paragens", []) or []) if "recebid" in self.desktop_main.norm_text(str(stop.get("pod_estado", "") or "")))
        zones = []
        for stop in list(detail.get("paragens", []) or []):
            zone_txt = str(stop.get("zona_transporte", "") or "").strip()
            if zone_txt and zone_txt not in zones:
                zones.append(zone_txt)
        detail["zonas"] = zones
        detail["vehicle_options"] = self._transport_vehicle_options()
        detail["driver_options"] = self._transport_driver_options()
        detail["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.ne_suppliers() or [])]
        detail["zone_options"] = self.transport_zone_options()
        return detail

    def transport_create_or_update(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip()
        trip = self._transport_find(numero) if numero else None
        if trip is None:
            numero = str(numero or self.desktop_main.next_transporte_numero(data)).strip()
            trip = {
                "numero": numero,
                "paragens": [],
                "created_by": str((self.user or {}).get("username", "") or "").strip(),
                "created_at": self.desktop_main.now_iso(),
            }
            data.setdefault("transportes", []).append(trip)
            self.desktop_main.reserve_transporte_numero(data, numero)
        defaults = self._transport_defaults()
        supplier_id, supplier_text, _supplier_contact = self._normalize_supplier_reference(
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
        trip["custo_previsto"] = round(self._parse_float(payload.get("custo_previsto", trip.get("custo_previsto", 0)), 0), 2)
        trip["paletes_total_manual"] = round(self._parse_float(payload.get("paletes_total_manual", trip.get("paletes_total_manual", 0)), 0), 2)
        trip["peso_total_manual_kg"] = round(self._parse_float(payload.get("peso_total_manual_kg", trip.get("peso_total_manual_kg", 0)), 0), 2)
        trip["volume_total_manual_m3"] = round(self._parse_float(payload.get("volume_total_manual_m3", trip.get("volume_total_manual_m3", 0)), 0), 3)
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
        if "subcontrat" in self.desktop_main.norm_text(trip["tipo_responsavel"]) and not trip["transportadora_nome"]:
            raise ValueError("Seleciona a transportadora externa para viagens subcontratadas.")
        trip["observacoes"] = str(payload.get("observacoes", trip.get("observacoes", "")) or "").strip()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_reindex_stops(trip)
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_request_service(self, numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).request_service(numero, payload))

    def transport_remove_trip(self, numero: str) -> None:
        trip_num = str(numero or "").strip()
        if not trip_num:
            raise ValueError("Seleciona uma viagem.")
        data = self.ensure_data()
        trips = [row for row in list(data.get("transportes", []) or []) if isinstance(row, dict)]
        target = next((row for row in trips if str(row.get("numero", "") or "").strip() == trip_num), None)
        if target is None:
            raise ValueError("Transporte nao encontrado.")
        data["transportes"] = [row for row in trips if row is not target]
        self._transport_sync_order_links()
        self._save(force=True)

    def transport_update_stop(self, numero: str, encomenda_numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).update(numero, encomenda_numero, payload))

    def transport_assign_orders(self, numero: str, order_numbers: list[str]) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        trip_state = str(trip.get("estado", "") or "Planeado").strip()
        if "conclu" in self.desktop_main.norm_text(trip_state) or "anulad" in self.desktop_main.norm_text(trip_state):
            raise ValueError("Nao podes alterar uma viagem concluida ou anulada.")
        active_assignments = {
            str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip(): str(tr.get("numero", "") or "").strip()
            for tr in list(self.ensure_data().get("transportes", []) or [])
            if isinstance(tr, dict) and "anulad" not in self.desktop_main.norm_text(str(tr.get("estado", "") or ""))
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
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        for enc_num in order_list:
            if enc_num in existing_orders:
                continue
            assigned_trip = active_assignments.get(enc_num)
            if assigned_trip and assigned_trip != trip.get("numero"):
                raise ValueError(f"A encomenda {enc_num} ja esta afeta ao transporte {assigned_trip}.")
            enc = self.get_encomenda_by_numero(enc_num)
            if enc is None:
                raise ValueError(f"Encomenda nao encontrada: {enc_num}")
            if not self._transport_is_own_cargo(enc):
                raise ValueError(f"A encomenda {enc_num} nao esta definida como transporte a nosso cargo.")
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = find_cliente_fn(self.ensure_data(), cli_code) if callable(find_cliente_fn) and cli_code else {}
            latest_guide = self._transport_latest_guide_for_order(enc_num) or {}
            metrics = self._transport_metrics_for_order(enc, cli_obj)
            carrier_id = str(trip.get("transportadora_id", "") or metrics.get("transportadora_id", "") or "").strip()
            carrier_name = str(trip.get("transportadora_nome", "") or metrics.get("transportadora_nome", "") or "").strip()
            zone_txt = str(metrics.get("zona_transporte", "") or "").strip()
            suggestion = self._transport_tariff_suggestion(
                carrier_id,
                carrier_name,
                zone_txt,
                metrics.get("paletes", 0.0),
                metrics.get("peso_bruto_kg", 0.0),
                metrics.get("volume_m3", 0.0),
            )
            order_cost = round(self._parse_float(metrics.get("custo_transporte", 0), 0), 2)
            suggested_cost = round(self._parse_float(suggestion.get("custo_sugerido", 0), 0), 2)
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
        self._transport_reindex_stops(trip)
        self._transport_sync_order_links()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_apply_suggested_cost(self, numero: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        detail = self.transport_detail(numero)
        suggested_map = {
            str(stop.get("encomenda_numero", "") or "").strip(): round(self._parse_float(stop.get("custo_sugerido", 0), 0), 2)
            for stop in list(detail.get("paragens", []) or [])
        }
        total = 0.0
        applied = 0
        for stop in list(trip.get("paragens", []) or []):
            if not isinstance(stop, dict):
                continue
            enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            suggested = round(self._parse_float(suggested_map.get(enc_num, 0), 0), 2)
            if suggested <= 0:
                continue
            stop["custo_transporte"] = suggested
            total += suggested
            applied += 1
        if applied <= 0:
            raise ValueError("Sem custos sugeridos para aplicar nesta viagem.")
        trip["custo_previsto"] = round(total, 2)
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_remove_stop(self, numero: str, encomenda_numero: str) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).remove(numero, encomenda_numero))

    def transport_move_stop(self, numero: str, encomenda_numero: str, direction: int) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).move(numero, encomenda_numero, direction))

    def transport_set_status(self, numero: str, estado: str) -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).set_status(numero, estado))

    def transport_set_stop_status(self, numero: str, encomenda_numero: str, estado: str, observacoes: str = "") -> dict[str, Any]:
        return self.transport_detail(transport_stops(self).set_stop_status(numero, encomenda_numero, estado, observacoes))

    def transport_route_sheet_render(self, numero: str, path: str | Path) -> Path:
        detail = self.transport_detail(numero)
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

        out_path = Path(path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page_w, page_h = A4
        margin = 32
        inner_w = page_w - (2 * margin)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_text = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_text) if logo_text and Path(logo_text).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        regular = "Helvetica"
        bold = "Helvetica-Bold"
        c = canvas.Canvas(str(out_path), pagesize=A4)
        c.setTitle(self._operator_pdf_text(f"Documento de viagem {detail.get('numero', '')}"))
        page_number = 1

        def text(value: Any) -> str:
            return self._operator_pdf_text(value)

        def draw_card(x: float, y: float, width: float, height: float, label: str, value: str, *, accent: bool = False) -> None:
            c.setFillColor(palette["primary_soft"] if accent else palette["surface"])
            c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
            c.roundRect(x, y, width, height, 3, stroke=1, fill=1)
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.2)
            c.drawString(x + 8, y + height - 11, text(_pdf_clip_text(label, width - 16, regular, 6.2)))
            value_font = 9.4
            while value_font > 6.5 and len(str(value or "")) * value_font * 0.52 > width - 16:
                value_font -= 0.3
            c.setFillColor(palette["ink"])
            c.setFont(bold, value_font)
            c.drawString(x + 8, y + 8, text(_pdf_clip_text(value or "-", width - 16, bold, value_font)))

        def draw_footer() -> None:
            c.setStrokeColor(palette["line"])
            c.line(margin, 27, page_w - margin, 27)
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.3)
            c.drawString(margin, 17, text(f"LUGEST | Documento de viagem {detail.get('numero', '-') or '-'}"))
            c.drawRightString(page_w - margin, 17, text(f"Pagina {page_number} | {printed_at}"))

        def draw_header() -> float:
            top = page_h - 26
            c.setFillColor(palette["surface"])
            c.rect(0, 0, page_w, page_h, stroke=0, fill=1)
            c.setFillColor(palette["primary"])
            c.rect(0, page_h - 9, page_w, 9, stroke=0, fill=1)
            if logo_path:
                self._draw_operator_logo_plate(c, palette, logo_path, margin, top - 54, 82, 40, radius=3, padding_x=4, padding_y=3)
            title_x = margin + 96
            title_w = inner_w - 190
            c.setFillColor(palette["primary_dark"])
            c.setFont(bold, 17)
            c.drawString(title_x, top - 21, text("Transportes | Documento de viagem"))
            c.setFillColor(palette["muted"])
            c.setFont(regular, 7.4)
            c.drawString(title_x, top - 37, text(_pdf_clip_text(f"Viagem {detail.get('numero', '-') or '-'} | {detail.get('data_planeada', '-') or '-'} as {detail.get('hora_saida', '-') or '-'}", title_w, regular, 7.4)))
            status = str(detail.get("estado", "") or "Planeado")
            draw_card(page_w - margin - 86, top - 52, 86, 38, "Estado", status, accent=True)

            metric_y = top - 101
            gap = 7
            metric_w = (inner_w - (3 * gap)) / 4
            metrics = [
                ("Paragens", str(len(list(detail.get("paragens", []) or [])))),
                ("Paletes", self._fmt(detail.get("paletes", 0))),
                ("Peso bruto", f"{self._fmt(detail.get('peso_bruto_kg', 0))} kg"),
                ("Volume", f"{self._fmt(detail.get('volume_m3', 0))} m3"),
            ]
            for index, (label, value) in enumerate(metrics):
                draw_card(margin + index * (metric_w + gap), metric_y, metric_w, 35, label, value, accent=index == 0)

            meta_y = metric_y - 66
            meta_gap = 8
            meta_w = (inner_w - meta_gap) / 2
            vehicle = " | ".join(
                part for part in (
                    str(detail.get("viatura", "") or "").strip(),
                    str(detail.get("matricula", "") or "").strip(),
                    str(detail.get("motorista", "") or "").strip(),
                    str(detail.get("telefone_motorista", "") or "").strip(),
                ) if part
            ) or "Por definir"
            carrier = " | ".join(
                part for part in (
                    str(detail.get("transportadora_nome", "") or "").strip(),
                    str(detail.get("referencia_transporte", "") or "").strip(),
                    str(detail.get("pedido_transporte_estado", "") or "").strip(),
                    str(detail.get("pedido_transporte_ref", "") or "").strip(),
                ) if part
            ) or "Transporte proprio / sem pedido externo"
            draw_card(margin, meta_y, meta_w, 50, "Viatura / matricula / motorista / contacto", vehicle)
            draw_card(margin + meta_w + meta_gap, meta_y, meta_w, 50, "Transportadora / referencia / pedido", carrier)
            return meta_y - 16

        columns = [
            ("Ord", 28), ("Encomenda", 74), ("Cliente", 90), ("Descarga", 138),
            ("Planeado", 72), ("Guia", 58), ("Estado", inner_w - 460),
        ]

        def draw_table_header(y: float) -> float:
            c.setFillColor(palette["primary_soft"])
            c.setStrokeColor(palette["line_strong"])
            c.rect(margin, y - 20, inner_w, 20, stroke=1, fill=1)
            c.setFillColor(palette["primary_dark"])
            c.setFont(bold, 6.8)
            x = margin
            for label, width in columns:
                c.drawString(x + 5, y - 13, text(label))
                x += width
            return y - 25

        def start_new_page() -> float:
            nonlocal page_number
            draw_footer()
            c.showPage()
            page_number += 1
            return draw_table_header(draw_header())

        y = draw_table_header(draw_header())
        for row_index, stop in enumerate(list(detail.get("paragens", []) or [])):
            checklist = (
                f"Carga {'OK' if stop.get('check_carga_ok') else '-'} | "
                f"Docs {'OK' if stop.get('check_docs_ok') else '-'} | "
                f"Paletes {'OK' if stop.get('check_paletes_ok') else '-'}"
            )
            notes = " | ".join(
                part for part in (
                    f"Zona {stop.get('zona_transporte', '-') or '-'}",
                    (
                        f"GPS {str(stop.get('latitude', '') or '').strip()},"
                        f"{str(stop.get('longitude', '') or '').strip()}"
                        if str(stop.get("latitude", "") or "").strip()
                        and str(stop.get("longitude", "") or "").strip()
                        else ""
                    ),
                    f"Carga {self._fmt(stop.get('paletes', 0))} pal / {self._fmt(stop.get('peso_bruto_kg', 0))} kg / {self._fmt(stop.get('volume_m3', 0))} m3",
                    checklist,
                    f"POD {stop.get('pod_estado', '-') or '-'}",
                    str(stop.get("observacoes", "") or "").strip(),
                ) if part
            )
            note_lines = _pdf_wrap_text(notes, regular, 6.4, inner_w - 18, max_lines=2) or []
            row_height = 27 + (len(note_lines) * 8)
            if y - row_height < 84:
                y = start_new_page()
            row_y = y - row_height
            c.setFillColor(palette["surface"] if row_index % 2 == 0 else palette["surface_alt"])
            c.setStrokeColor(palette["line"])
            c.rect(margin, row_y, inner_w, row_height - 3, stroke=1, fill=1)
            values = [
                str(stop.get("ordem", "-") or "-"),
                str(stop.get("encomenda_numero", "-") or "-"),
                str(stop.get("cliente_nome", "-") or "-"),
                str(stop.get("local_descarga", "-") or "-"),
                str(stop.get("data_planeada", "") or detail.get("data_planeada", "-")).replace("T", " ")[:16] or "-",
                str(stop.get("guia_numero", "-") or "-"),
                str(stop.get("estado", "-") or "-"),
            ]
            x = margin
            for column_index, (value, (_label, width)) in enumerate(zip(values, columns)):
                font_name = bold if column_index in (0, 1, 6) else regular
                c.setFillColor(palette["ink"] if column_index in (0, 1, 6) else palette["muted"])
                c.setFont(font_name, 6.7)
                c.drawString(x + 5, y - 17, text(_pdf_clip_text(value, width - 10, font_name, 6.7)))
                x += width
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.4)
            note_y = y - 27
            for line in note_lines:
                c.drawString(margin + 9, note_y, text(line))
                note_y -= 8
            y = row_y - 5

        if y < 92:
            y = start_new_page()
        signature_y = 50
        c.setStrokeColor(palette["line_strong"])
        signature_w = (inner_w - 20) / 3
        for index, label in enumerate(("Motorista / saida", "Conferencia de carga", "Rececao / chegada")):
            x = margin + index * (signature_w + 10)
            c.line(x, signature_y + 16, x + signature_w, signature_y + 16)
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.2)
            c.drawCentredString(x + signature_w / 2, signature_y + 6, text(label))
        draw_footer()
        c.save()
        return out_path

    def _transport_route_sheet_render_legacy(self, numero: str, path: str | Path) -> Path:
        detail = self.transport_detail(numero)
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

        out_path = Path(path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page_w, page_h = A4
        margin = 34
        row_h = 22
        c = canvas.Canvas(str(out_path), pagesize=A4)

        def draw_header() -> float:
            c.setTitle(f"Folha de rota {detail.get('numero', '')}")
            c.setFont("Helvetica-Bold", 20)
            c.setFillColor(colors.HexColor("#0f172a"))
            c.drawString(margin, page_h - 44, "Transportes | Folha de rota")
            c.setFont("Helvetica-Bold", 12)
            c.drawString(margin, page_h - 64, f"Viagem {detail.get('numero', '-')}")
            c.setFont("Helvetica", 9)
            c.setFillColor(colors.HexColor("#475569"))
            meta = [
                f"Data {detail.get('data_planeada', '-') or '-'}",
                f"Saida {detail.get('hora_saida', '-') or '-'}",
                f"Tipo {detail.get('tipo_responsavel', '-') or '-'}",
                f"Estado {detail.get('estado', '-') or '-'}",
                f"Viatura {detail.get('viatura', '-') or '-'}",
                f"Motorista {detail.get('motorista', '-') or '-'}",
            ]
            c.drawString(margin, page_h - 80, " | ".join(meta))
            carrier_txt = str(detail.get("transportadora_nome", "") or "-").strip() or "-"
            c.drawString(
                margin,
                page_h - 94,
                f"Origem {detail.get('origem', '-') or '-'} | Transportadora {carrier_txt} | Ref {detail.get('referencia_transporte', '-') or '-'}",
            )
            c.drawString(
                margin,
                page_h - 108,
                f"Totais {detail.get('paletes', 0):.2f} pal | {detail.get('peso_bruto_kg', 0):.1f} kg | "
                f"{detail.get('volume_m3', 0):.3f} m3 | Preco {self._fmt_eur(detail.get('preco_total', 0))} | "
                f"Custo {self._fmt_eur(detail.get('custo_total', 0))} | Sug. {self._fmt_eur(detail.get('custo_sugerido_total', 0))}",
            )
            c.drawString(
                margin,
                page_h - 122,
                f"Pedido transporte {detail.get('pedido_transporte_estado', 'Nao pedido') or 'Nao pedido'} | "
                f"Ref pedido {detail.get('pedido_transporte_ref', '-') or '-'}",
            )
            response_parts = []
            if detail.get("pedido_confirmado_at"):
                response_parts.append(f"Confirmado {detail.get('pedido_confirmado_at', '-')}")
            if detail.get("pedido_recusado_at"):
                response_parts.append(f"Recusado {detail.get('pedido_recusado_at', '-')}")
            if detail.get("pedido_resposta_obs"):
                response_parts.append(f"Resposta {detail.get('pedido_resposta_obs', '-')}")
            if response_parts:
                c.drawString(margin, page_h - 136, " | ".join(response_parts))
                line_y = page_h - 146
            else:
                line_y = page_h - 132
            c.setStrokeColor(colors.HexColor("#cbd5e1"))
            c.line(margin, line_y, page_w - margin, line_y)
            return line_y - 18

        def draw_table_header(y: float) -> float:
            c.setFillColor(colors.HexColor("#0f172a"))
            c.roundRect(margin, y - row_h + 4, page_w - (margin * 2), row_h, 8, fill=1, stroke=0)
            cols = [("Ord", 34), ("Encomenda", 84), ("Cliente", 120), ("Descarga", 168), ("Planeado", 82), ("Guia", 64), ("Estado", 74)]
            x = margin + 8
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 8)
            for label, width in cols:
                c.drawString(x, y - 10, label)
                x += width
            return y - row_h - 2

        def new_page() -> float:
            c.showPage()
            return draw_header()

        y = draw_header()
        y = draw_table_header(y)
        widths = [34, 84, 120, 168, 82, 64, 74]
        for stop in list(detail.get("paragens", []) or []):
            metrics_line = (
                f"Pal {self._fmt(stop.get('paletes', 0))} | "
                f"{self._fmt(stop.get('peso_bruto_kg', 0))} kg | "
                f"{self._fmt(stop.get('volume_m3', 0))} m3 | "
                f"Preco {self._fmt_eur(stop.get('preco_transporte', 0))} | "
                f"Custo {self._fmt_eur(stop.get('custo_transporte', 0))} | Sug. {self._fmt_eur(stop.get('custo_sugerido', 0))}"
            )
            carrier_line = ""
            if stop.get("transportadora_nome"):
                carrier_line = f"Transportadora: {stop.get('transportadora_nome', '-')}"
                if stop.get("referencia_transporte"):
                    carrier_line += f" | Ref: {stop.get('referencia_transporte', '-')}"
            zone_line = ""
            if stop.get("zona_transporte"):
                zone_line = f"Zona: {stop.get('zona_transporte', '-')}"
                if stop.get("tarifario_label"):
                    zone_line += f" | Tarifario: {stop.get('tarifario_label', '-')}"
            checklist_line = (
                f"Checklist: carga {'OK' if stop.get('check_carga_ok') else '-'} / "
                f"docs {'OK' if stop.get('check_docs_ok') else '-'} / "
                f"paletes {'OK' if stop.get('check_paletes_ok') else '-'}"
            )
            pod_line = ""
            if stop.get("pod_estado"):
                pod_line = f"POD: {stop.get('pod_estado', '-')}"
                if stop.get("pod_recebido_nome"):
                    pod_line += f" por {stop.get('pod_recebido_nome', '-')}"
            combined_note = " | ".join(
                [
                    part
                    for part in [
                        metrics_line,
                        carrier_line,
                        zone_line,
                        checklist_line,
                        pod_line,
                        str(stop.get("pod_obs", "") or "").strip(),
                        str(stop.get("observacoes", "") or "").strip(),
                    ]
                    if part
                ]
            )
            extra_lines = _pdf_wrap_text(combined_note, "Helvetica", 7.0, page_w - (margin * 2) - 16, max_lines=3)
            needed = row_h + (8 * len(extra_lines)) + 8
            if y < margin + needed:
                y = new_page()
                y = draw_table_header(y)
            c.setFillColor(colors.HexColor("#f8fafc"))
            c.roundRect(margin, y - row_h + 4, page_w - (margin * 2), row_h, 6, fill=1, stroke=0)
            values = [
                str(stop.get("ordem", "-") or "-"),
                _pdf_clip_text(stop.get("encomenda_numero", "-"), widths[1] - 6, "Helvetica-Bold", 7.6),
                _pdf_clip_text(stop.get("cliente_nome", "-"), widths[2] - 6, "Helvetica", 7.4),
                _pdf_clip_text(stop.get("local_descarga", "-"), widths[3] - 6, "Helvetica", 7.2),
                _pdf_clip_text(str(stop.get("data_planeada", "") or detail.get("data_planeada", "-")).replace("T", " ")[:16] or "-", widths[4] - 6, "Helvetica", 7.4),
                _pdf_clip_text(stop.get("guia_numero", "-"), widths[5] - 6, "Helvetica", 7.4),
                _pdf_clip_text(stop.get("estado", "-"), widths[6] - 6, "Helvetica-Bold", 7.4),
            ]
            x = margin + 8
            c.setFillColor(colors.HexColor("#0f172a"))
            for index, value in enumerate(values):
                c.setFont("Helvetica-Bold" if index in (0, 1, 6) else "Helvetica", 7.4)
                c.drawString(x, y - 10, str(value or "-"))
                x += widths[index]
            if extra_lines:
                c.setFillColor(colors.HexColor("#64748b"))
                c.setFont("Helvetica", 7.0)
                text_y = y - 19
                for line in extra_lines:
                    c.drawString(margin + 12, text_y, line)
                    text_y -= 8
                y = text_y - 6
            else:
                y -= row_h + 4
        if y < 110:
            y = new_page()
        c.setStrokeColor(colors.HexColor("#cbd5e1"))
        c.line(margin, 92, page_w - margin, 92)
        c.setFont("Helvetica", 8)
        c.setFillColor(colors.HexColor("#475569"))
        c.drawString(margin, 76, "Observacao: esta folha de rota apoia a distribuicao e nao substitui a guia/documento de transporte.")
        c.drawString(margin, 58, "Motorista: ____________________________")
        c.drawString(margin + 220, 58, "Saida: ____________")
        c.drawString(margin + 360, 58, "Chegada: ____________")
        c.save()
        return out_path

    def transport_route_sheet_open(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_transporte_{str(numero or '').strip()}.pdf"
        self.transport_route_sheet_render(numero, target)
        os.startfile(str(target))
        return target

