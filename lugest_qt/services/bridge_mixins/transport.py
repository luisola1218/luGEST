from __future__ import annotations

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
        return {
            "id": "",
            "transportadora_id": "",
            "transportadora_nome": "",
            "zona": "",
            "valor_base": 0.0,
            "valor_por_palete": 0.0,
            "valor_por_kg": 0.0,
            "valor_por_m3": 0.0,
            "custo_minimo": 0.0,
            "ativo": True,
            "observacoes": "",
        }

    def _transport_tariff_signature(self, row: dict[str, Any] | None) -> str:
        row = dict(row or {})
        carrier_parts = [
            part
            for part in [
                str(row.get("transportadora_id", "") or "").strip(),
                str(row.get("transportadora_nome", "") or "").strip(),
            ]
            if part
        ]
        carrier = " - ".join(carrier_parts).strip(" -") or "Sem transportadora"
        zone = str(row.get("zona", "") or "").strip() or "Sem zona"
        return f"{carrier} | {zone}"

    def _transport_tariff_next_id(self) -> int:
        highest = 0
        for row in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            highest = max(highest, int(self._parse_float((row or {}).get("id", 0), 0) or 0))
        return highest + 1

    def _transport_tariff_match(self, transportadora_id: Any = "", transportadora_nome: Any = "", zona: Any = "") -> dict[str, Any] | None:
        zone_norm = self.desktop_main.norm_text(zona)
        if not zone_norm:
            return None
        supplier_id = str(transportadora_id or "").strip()
        supplier_name_norm = self.desktop_main.norm_text(transportadora_nome)
        best: tuple[int, int, dict[str, Any]] | None = None
        for raw in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            if not isinstance(raw, dict) or not bool(raw.get("ativo", True)):
                continue
            if self.desktop_main.norm_text(raw.get("zona", "")) != zone_norm:
                continue
            row_supplier_id = str(raw.get("transportadora_id", "") or "").strip()
            row_supplier_name_norm = self.desktop_main.norm_text(raw.get("transportadora_nome", ""))
            score = 0
            if supplier_id and row_supplier_id and row_supplier_id == supplier_id:
                score = 3
            elif supplier_name_norm and row_supplier_name_norm and row_supplier_name_norm == supplier_name_norm:
                score = 2
            elif not row_supplier_id and not row_supplier_name_norm:
                score = 1
            if score <= 0:
                continue
            row_id = int(self._parse_float(raw.get("id", 0), 0) or 0)
            candidate = (score, -row_id, raw)
            if best is None or candidate > best:
                best = candidate
        return dict(best[2]) if best else None

    def _transport_tariff_cost_from_row(self, row: dict[str, Any] | None, paletes: Any = 0, peso_bruto_kg: Any = 0, volume_m3: Any = 0) -> float:
        row = dict(row or {})
        base = round(self._parse_float(row.get("valor_base", 0), 0), 2)
        per_pal = round(self._parse_float(row.get("valor_por_palete", 0), 0), 2)
        per_kg = round(self._parse_float(row.get("valor_por_kg", 0), 0), 4)
        per_m3 = round(self._parse_float(row.get("valor_por_m3", 0), 0), 2)
        minimum = round(self._parse_float(row.get("custo_minimo", 0), 0), 2)
        pal = max(0.0, round(self._parse_float(paletes, 0), 2))
        peso = max(0.0, round(self._parse_float(peso_bruto_kg, 0), 2))
        volume = max(0.0, round(self._parse_float(volume_m3, 0), 3))
        total = round(base + (pal * per_pal) + (peso * per_kg) + (volume * per_m3), 2)
        if minimum > 0:
            total = max(total, minimum)
        return round(total, 2)

    def _transport_tariff_suggestion(
        self,
        transportadora_id: Any = "",
        transportadora_nome: Any = "",
        zona: Any = "",
        paletes: Any = 0,
        peso_bruto_kg: Any = 0,
        volume_m3: Any = 0,
    ) -> dict[str, Any]:
        tariff = self._transport_tariff_match(transportadora_id, transportadora_nome, zona)
        if tariff is None:
            return {
                "tarifario_id": "",
                "tarifario_label": "",
                "custo_sugerido": 0.0,
            }
        return {
            "tarifario_id": tariff.get("id", ""),
            "tarifario_label": self._transport_tariff_signature(tariff),
            "custo_sugerido": self._transport_tariff_cost_from_row(tariff, paletes, peso_bruto_kg, volume_m3),
        }

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

    def _transport_reindex_stops(self, trip: dict[str, Any]) -> None:
        stops = list(trip.get("paragens", []) or [])
        stops.sort(
            key=lambda row: (
                int(self._parse_float((row or {}).get("ordem", 0), 0) or 0),
                str((row or {}).get("data_planeada", "") or ""),
                str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or ""),
            )
        )
        for index, stop in enumerate(stops, start=1):
            stop["ordem"] = index
        trip["paragens"] = stops

    def _transport_sync_order_links(self) -> None:
        assigned: dict[str, tuple[str, str]] = {}
        for tr in list(self.ensure_data().get("transportes", []) or []):
            if not isinstance(tr, dict):
                continue
            trip_num = str(tr.get("numero", "") or "").strip()
            trip_state = str(tr.get("estado", "") or "").strip() or "Planeado"
            if not trip_num or "anulad" in self.desktop_main.norm_text(trip_state):
                continue
            for stop in list(tr.get("paragens", []) or []):
                if not isinstance(stop, dict):
                    continue
                enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
                if not enc_num:
                    continue
                assigned[enc_num] = (trip_num, self._transport_stop_state(stop, trip_state))
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            if not isinstance(enc, dict):
                continue
            num = str(enc.get("numero", "") or "").strip()
            if not num:
                continue
            trip_info = assigned.get(num)
            if trip_info is None:
                enc["transporte_numero"] = ""
                enc["estado_transporte"] = ""
                continue
            enc["transporte_numero"] = trip_info[0]
            enc["estado_transporte"] = trip_info[1]

    def transport_defaults(self) -> dict[str, Any]:
        payload = dict(self._transport_defaults())
        payload["vehicle_options"] = self._transport_vehicle_options()
        payload["driver_options"] = self._transport_driver_options()
        payload["supplier_options"] = [f"{row.get('id', '')} - {row.get('nome', '')}".strip(" -") for row in list(self.ne_suppliers() or [])]
        payload["zone_options"] = self.transport_zone_options()
        return payload

    def transport_tariff_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("transportes_tarifarios", []) or []):
            if not isinstance(raw, dict):
                continue
            row = {
                "id": int(self._parse_float(raw.get("id", 0), 0) or 0),
                "transportadora_id": str(raw.get("transportadora_id", "") or "").strip(),
                "transportadora_nome": str(raw.get("transportadora_nome", "") or "").strip(),
                "zona": str(raw.get("zona", "") or "").strip(),
                "valor_base": round(self._parse_float(raw.get("valor_base", 0), 0), 2),
                "valor_por_palete": round(self._parse_float(raw.get("valor_por_palete", 0), 0), 2),
                "valor_por_kg": round(self._parse_float(raw.get("valor_por_kg", 0), 0), 4),
                "valor_por_m3": round(self._parse_float(raw.get("valor_por_m3", 0), 0), 2),
                "custo_minimo": round(self._parse_float(raw.get("custo_minimo", 0), 0), 2),
                "ativo": bool(raw.get("ativo", True)),
                "observacoes": str(raw.get("observacoes", "") or "").strip(),
                "label": self._transport_tariff_signature(raw),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (self.desktop_main.norm_text(item.get("transportadora_nome", "")), self.desktop_main.norm_text(item.get("zona", "")), item.get("id", 0)))
        return rows

    def transport_tariff_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        rows = data.setdefault("transportes_tarifarios", [])
        tariff_id = int(self._parse_float(payload.get("id", 0), 0) or 0)
        zone = str(payload.get("zona", "") or "").strip()
        if not zone:
            raise ValueError("Zona obrigatoria no tarifario.")
        transportadora_id, transportadora_nome, _contact = self._normalize_supplier_reference(
            payload.get("transportadora_id", ""),
            payload.get("transportadora_nome", ""),
        )
        zone_norm = self.desktop_main.norm_text(zone)
        for row in rows:
            if not isinstance(row, dict):
                continue
            if int(self._parse_float(row.get("id", 0), 0) or 0) == tariff_id:
                continue
            same_zone = self.desktop_main.norm_text(row.get("zona", "")) == zone_norm
            same_supplier = (
                str(row.get("transportadora_id", "") or "").strip() == transportadora_id
                and self.desktop_main.norm_text(row.get("transportadora_nome", "")) == self.desktop_main.norm_text(transportadora_nome)
            )
            if same_zone and same_supplier:
                raise ValueError("Ja existe um tarifario para essa transportadora e zona.")
        target = next((row for row in rows if int(self._parse_float((row or {}).get("id", 0), 0) or 0) == tariff_id), None) if tariff_id > 0 else None
        if target is None:
            target = self.transport_tariff_defaults()
            target["id"] = self._transport_tariff_next_id()
            rows.append(target)
        target["transportadora_id"] = transportadora_id
        target["transportadora_nome"] = transportadora_nome
        target["zona"] = zone
        target["valor_base"] = round(self._parse_float(payload.get("valor_base", target.get("valor_base", 0)), 0), 2)
        target["valor_por_palete"] = round(self._parse_float(payload.get("valor_por_palete", target.get("valor_por_palete", 0)), 0), 2)
        target["valor_por_kg"] = round(self._parse_float(payload.get("valor_por_kg", target.get("valor_por_kg", 0)), 0), 4)
        target["valor_por_m3"] = round(self._parse_float(payload.get("valor_por_m3", target.get("valor_por_m3", 0)), 0), 2)
        target["custo_minimo"] = round(self._parse_float(payload.get("custo_minimo", target.get("custo_minimo", 0)), 0), 2)
        target["ativo"] = bool(payload.get("ativo", target.get("ativo", True)))
        target["observacoes"] = str(payload.get("observacoes", target.get("observacoes", "")) or "").strip()
        self._save(force=True)
        return dict(target)

    def transport_tariff_remove(self, tariff_id: Any) -> None:
        target_id = int(self._parse_float(tariff_id, 0), 0)
        rows = list(self.ensure_data().get("transportes_tarifarios", []) or [])
        filtered = [row for row in rows if int(self._parse_float((row or {}).get("id", 0), 0) or 0) != target_id]
        if len(filtered) == len(rows):
            raise ValueError("Tarifario nao encontrado.")
        self.ensure_data()["transportes_tarifarios"] = filtered
        self._save(force=True)

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
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        payload = dict(payload or {})
        supplier_id, supplier_text, _supplier_contact = self._normalize_supplier_reference(
            payload.get("transportadora_id", trip.get("transportadora_id", "")),
            payload.get("transportadora_nome", trip.get("transportadora_nome", "")),
        )
        if "transportadora_id" in payload or "transportadora_nome" in payload:
            trip["transportadora_id"] = supplier_id
            trip["transportadora_nome"] = supplier_text
        trip["paletes_total_manual"] = round(self._parse_float(payload.get("paletes_total_manual", trip.get("paletes_total_manual", 0)), 0), 2)
        trip["peso_total_manual_kg"] = round(self._parse_float(payload.get("peso_total_manual_kg", trip.get("peso_total_manual_kg", 0)), 0), 2)
        trip["volume_total_manual_m3"] = round(self._parse_float(payload.get("volume_total_manual_m3", trip.get("volume_total_manual_m3", 0)), 0), 3)
        trip["custo_previsto"] = round(self._parse_float(payload.get("custo_previsto", trip.get("custo_previsto", 0)), 0), 2)
        request_state = str(payload.get("pedido_transporte_estado", trip.get("pedido_transporte_estado", "Pedido enviado")) or "Pedido enviado").strip() or "Pedido enviado"
        trip["pedido_transporte_estado"] = request_state
        trip["pedido_transporte_ref"] = str(payload.get("pedido_transporte_ref", trip.get("pedido_transporte_ref", "")) or "").strip()
        trip["pedido_transporte_obs"] = str(payload.get("pedido_transporte_obs", trip.get("pedido_transporte_obs", "")) or "").strip()
        trip["pedido_resposta_obs"] = str(payload.get("pedido_resposta_obs", trip.get("pedido_resposta_obs", "")) or "").strip()
        normalized_state = self.desktop_main.norm_text(request_state)
        if normalized_state in {"nao pedido", "nao-pedido"}:
            trip["pedido_transporte_at"] = ""
            trip["pedido_transporte_by"] = ""
            trip["pedido_confirmado_at"] = ""
            trip["pedido_confirmado_by"] = ""
            trip["pedido_recusado_at"] = ""
            trip["pedido_recusado_by"] = ""
        else:
            trip["pedido_transporte_at"] = self.desktop_main.now_iso()
            trip["pedido_transporte_by"] = str((self.user or {}).get("username", "") or "").strip()
            if "confirm" in normalized_state:
                trip["pedido_confirmado_at"] = self.desktop_main.now_iso()
                trip["pedido_confirmado_by"] = str((self.user or {}).get("username", "") or "").strip()
                trip["pedido_recusado_at"] = ""
                trip["pedido_recusado_by"] = ""
            elif "recus" in normalized_state:
                trip["pedido_recusado_at"] = self.desktop_main.now_iso()
                trip["pedido_recusado_by"] = str((self.user or {}).get("username", "") or "").strip()
                trip["pedido_confirmado_at"] = ""
                trip["pedido_confirmado_by"] = ""
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

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
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
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
            valid_guides = {row.get("numero", "") for row in self.transport_guide_options(enc_num)}
            if valid_guides and guide_number not in valid_guides:
                raise ValueError("A guia escolhida nao pertence a esta encomenda.")
        target["expedicao_numero"] = guide_number
        target["zona_transporte"] = str(payload.get("zona_transporte", target.get("zona_transporte", "")) or "").strip()
        target["local_descarga"] = str(payload.get("local_descarga", target.get("local_descarga", "")) or "").strip()
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
            "recebid" in self.desktop_main.norm_text(str(target.get("pod_estado", "") or ""))
            and not str(target.get("pod_recebido_at", "") or "").strip()
        ):
            target["pod_recebido_at"] = self.desktop_main.now_iso()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

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
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        enc_num = str(encomenda_numero or "").strip()
        before = len(list(trip.get("paragens", []) or []))
        trip["paragens"] = [
            row
            for row in list(trip.get("paragens", []) or [])
            if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() != enc_num
        ]
        if len(list(trip.get("paragens", []) or [])) == before:
            raise ValueError("Paragem nao encontrada.")
        self._transport_reindex_stops(trip)
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_move_stop(self, numero: str, encomenda_numero: str, direction: int) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
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
            return self.transport_detail(numero)
        stops[index], stops[target] = stops[target], stops[index]
        trip["paragens"] = stops
        self._transport_reindex_stops(trip)
        trip["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_set_status(self, numero: str, estado: str) -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
        state_txt = str(estado or "").strip()
        if not state_txt:
            raise ValueError("Estado obrigatorio.")
        trip["estado"] = state_txt
        if "conclu" in self.desktop_main.norm_text(state_txt):
            for stop in list(trip.get("paragens", []) or []):
                if not isinstance(stop, dict):
                    continue
                if "inciden" in self.desktop_main.norm_text(str(stop.get("estado", "") or "")):
                    continue
                stop["estado"] = "Entregue"
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_set_stop_status(self, numero: str, encomenda_numero: str, estado: str, observacoes: str = "") -> dict[str, Any]:
        trip = self._transport_find(numero)
        if trip is None:
            raise ValueError("Transporte nao encontrado.")
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
        target["estado"] = str(estado or "").strip() or "Planeada"
        if str(observacoes or "").strip():
            target["observacoes"] = str(observacoes or "").strip()
        trip["updated_at"] = self.desktop_main.now_iso()
        self._transport_sync_order_links()
        self._save(force=True)
        return self.transport_detail(numero)

    def transport_route_sheet_render(self, numero: str, path: str | Path) -> Path:
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

