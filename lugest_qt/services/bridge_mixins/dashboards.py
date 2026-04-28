from __future__ import annotations

from datetime import date, datetime
from typing import Any


class DashboardBridgeMixin:
    """Finance, operational, and summary dashboards for the Qt bridge."""

    def finance_dashboard(self, ano: str = "Todos") -> dict[str, Any]:
        data = self.ensure_data()
        year_filter = str(ano or "Todos").strip()
        all_years: set[str] = set()
        valor_produtos = 0.0
        compras_produtos_total = 0.0
        compras_materias_total = 0.0
        compras_fornecedor_totais: dict[str, float] = {}
        compras_mes_totais: dict[str, float] = {}
        produtos_rows = []
        for prod in list(data.get("produtos", []) or []):
            qty = self._parse_float(prod.get("qty", 0), 0)
            if qty <= 0:
                continue
            price_unit = self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0)
            total = round(qty * price_unit, 2)
            valor_produtos += total
            produtos_rows.append(
                {
                    "codigo": str(prod.get("codigo", "") or "").strip(),
                    "descricao": str(prod.get("descricao", "") or "").strip(),
                    "qty": round(qty, 2),
                    "preco_unid": round(price_unit, 4),
                    "valor": total,
                }
            )
        valor_materias = 0.0
        materias_rows = []
        for mat in list(data.get("materiais", []) or []):
            qty = self._parse_float(mat.get("quantidade", 0), 0)
            if qty <= 0:
                continue
            price_unit = self._parse_float(self.materia_actions._materia_preco_unid_record(mat), 0)
            total = round(qty * price_unit, 2)
            valor_materias += total
            materias_rows.append(
                {
                    "id": str(mat.get("id", "") or "").strip(),
                    "material": str(mat.get("material", "") or "").strip(),
                    "espessura": str(mat.get("espessura", "") or "").strip(),
                    "qty": round(qty, 2),
                    "preco_unid": round(price_unit, 4),
                    "valor": total,
                }
            )
        compras_rows = []
        compras_materias_rows = []
        compras_produtos_rows = []
        valor_ne_aprovadas = 0.0
        for note in list(data.get("notas_encomenda", []) or []):
            estado_txt = str(note.get("estado", "") or "").strip()
            estado_norm = self.desktop_main.norm_text(estado_txt)
            note_date = str(note.get("data_entrega", "") or note.get("data_documento", "") or "").strip()
            note_year = note_date[:4] if len(note_date) >= 4 and note_date[:4].isdigit() else ""
            if note_year:
                all_years.add(note_year)
            total_note = round(self._parse_float(note.get("total", 0), 0), 2)
            if "aprov" in estado_norm and (year_filter.lower() in ("todos", "todas", "all", "") or note_year == year_filter):
                valor_ne_aprovadas += total_note
            for line in list(note.get("linhas", []) or []):
                qtd_tot = self._parse_float(line.get("qtd", 0), 0)
                qtd_ent = self._parse_float(line.get("qtd_entregue", 0), 0)
                entregue = bool(line.get("entregue") or line.get("_stock_in"))
                qtd_hist = qtd_ent if qtd_ent > 0 else (qtd_tot if entregue else 0.0)
                if qtd_hist <= 0:
                    continue
                line_date = str(line.get("data_doc_entrega", "") or line.get("data_entrega_real", "") or note_date or "").strip()
                line_year = line_date[:4] if len(line_date) >= 4 and line_date[:4].isdigit() else ""
                if line_year:
                    all_years.add(line_year)
                if year_filter.lower() not in ("todos", "todas", "all", "") and line_year != year_filter:
                    continue
                preco = self._parse_float(line.get("preco", 0), 0)
                total_l = round(qtd_hist * preco, 2)
                compras_rows.append(
                    {
                        "data": line_date,
                        "ne": str(note.get("numero", "") or "").strip(),
                        "fornecedor": str(note.get("fornecedor", "") or "").strip(),
                        "artigo": str(line.get("descricao", "") or line.get("ref", "") or "").strip(),
                        "qtd": round(qtd_hist, 2),
                        "preco": round(preco, 4),
                        "total": total_l,
                        "estado": estado_txt,
                        "origem": str(line.get("origem", "") or "").strip(),
                    }
                )
                fornecedor_nome = str(note.get("fornecedor", "") or "").strip() or "Sem fornecedor"
                compras_fornecedor_totais[fornecedor_nome] = round(compras_fornecedor_totais.get(fornecedor_nome, 0.0) + total_l, 2)
                mes_key = "-"
                if len(line_date) >= 7:
                    mes_key = line_date[:7]
                compras_mes_totais[mes_key] = round(compras_mes_totais.get(mes_key, 0.0) + total_l, 2)
                origem_norm = self.desktop_main.norm_text(line.get("origem", ""))
                if "mater" in origem_norm:
                    compras_materias_total += total_l
                    compras_materias_rows.append(dict(compras_rows[-1]))
                else:
                    compras_produtos_total += total_l
                    compras_produtos_rows.append(dict(compras_rows[-1]))
        status_counts = {"Preparacao": 0, "Montagem": 0, "Em producao": 0, "Em pausa": 0, "Avaria": 0, "Concluida": 0}
        montagem_alertas = []
        for enc in list(data.get("encomendas", []) or []):
            estado = str(enc.get("estado", "") or "").strip()
            norm = self.desktop_main.norm_text(estado)
            montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "").strip()
            shortages = self._order_montagem_shortages(enc)
            if montagem_estado == "Pendente":
                cliente_codigo = str(enc.get("cliente", "") or "").strip()
                cliente_nome = next(
                    (
                        str(row.get("nome", "") or "").strip()
                        for row in list(data.get("clientes", []) or [])
                        if str(row.get("codigo", "") or "").strip() == cliente_codigo
                    ),
                    "",
                )
                stock_txt = "Stock OK"
                shortage_total = round(sum(self._parse_float(row.get("qtd_em_falta", 0), 0) for row in shortages), 2)
                supplier_options: list[str] = []
                for shortage in shortages:
                    supplier_txt = str(shortage.get("fornecedor_sugerido", "") or "").strip()
                    if supplier_txt and supplier_txt.lower() not in [value.lower() for value in supplier_options]:
                        supplier_options.append(supplier_txt)
                supplier_txt = "Por validar"
                if supplier_options:
                    supplier_txt = supplier_options[0]
                    if len(supplier_options) > 1:
                        supplier_txt += f" +{len(supplier_options) - 1}"
                if shortages:
                    first = shortages[0]
                    stock_txt = f"Falta {self._fmt(first.get('qtd_em_falta', 0))} {first.get('produto_codigo', '-')}"
                    if len(shortages) > 1:
                        stock_txt += f" +{len(shortages) - 1}"
                montagem_alertas.append(
                    {
                        "numero": str(enc.get("numero", "") or "").strip(),
                        "cliente": " - ".join([x for x in [cliente_codigo, cliente_nome] if x]).strip(" -"),
                        "montagem": str(self.desktop_main.encomenda_montagem_resumo(enc) or "Montagem final").strip(),
                        "tempo_min": round(self._parse_float(self.desktop_main.encomenda_montagem_tempo_min(enc), 0), 1),
                        "qtd_falta": shortage_total,
                        "fornecedor": supplier_txt if shortages else "-",
                        "stock": stock_txt,
                        "shortage_count": len(shortages),
                        "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                    }
                )
            if "avari" in norm:
                status_counts["Avaria"] += 1
            elif "montag" in norm:
                status_counts["Montagem"] += 1
            elif "produc" in norm or "curso" in norm:
                status_counts["Em producao"] += 1
            elif "paus" in norm or "interromp" in norm:
                status_counts["Em pausa"] += 1
            elif "concl" in norm:
                status_counts["Concluida"] += 1
            else:
                status_counts["Preparacao"] += 1
        montagem_alertas.sort(
            key=lambda item: (
                -int(item.get("shortage_count", 0) or 0),
                str(item.get("data_entrega", "") or "9999-99-99"),
                str(item.get("numero", "") or ""),
            )
        )
        produtos_rows.sort(key=lambda item: item.get("valor", 0), reverse=True)
        materias_rows.sort(key=lambda item: item.get("valor", 0), reverse=True)
        compras_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        compras_materias_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        compras_produtos_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        compras_fornecedor_rows = sorted(
            [{"fornecedor": key, "total": value} for key, value in compras_fornecedor_totais.items()],
            key=lambda item: item.get("total", 0),
            reverse=True,
        )
        compras_mes_rows = sorted(
            [{"mes": key, "total": value} for key, value in compras_mes_totais.items()],
            key=lambda item: str(item.get("mes", "") or ""),
            reverse=True,
        )
        subtitle_suffix = f"Ano {year_filter}" if year_filter.lower() not in ("todos", "todas", "all", "") else "Todos os anos"
        return {
            "cards": [
                {"title": "Stock MP", "value": self._fmt_eur(valor_materias), "subtitle": f"{len(data.get('materiais', []))} referencias", "tone": "warning"},
                {"title": "Stock Produtos", "value": self._fmt_eur(valor_produtos), "subtitle": f"{len(data.get('produtos', []))} refs | montagem {len(montagem_alertas)}", "tone": "success"},
                {"title": "Compras MP", "value": self._fmt_eur(compras_materias_total), "subtitle": subtitle_suffix, "tone": "warning"},
                {"title": "Compras Produtos", "value": self._fmt_eur(compras_produtos_total), "subtitle": subtitle_suffix, "tone": "success"},
                {"title": "Stock Total", "value": self._fmt_eur(valor_produtos + valor_materias), "subtitle": "Matéria-prima + produto acabado", "tone": "info"},
                {"title": "NE Aprovadas", "value": self._fmt_eur(valor_ne_aprovadas), "subtitle": subtitle_suffix, "tone": "default"},
            ],
            "order_status": [{"estado": key, "total": value} for key, value in status_counts.items()],
            "top_materias": materias_rows[:10],
            "top_produtos": produtos_rows[:10],
            "compras": compras_rows[:18],
            "compras_materias": compras_materias_rows[:18],
            "compras_produtos": compras_produtos_rows[:18],
            "compras_por_fornecedor": compras_fornecedor_rows[:12],
            "compras_por_mes": compras_mes_rows[:12],
            "montagem_alertas": montagem_alertas[:12],
            "years": sorted(all_years, reverse=True),
            "selected_year": year_filter or "Todos",
        }

    def operational_dashboard(self, ano: str = "Todos") -> dict[str, Any]:
        data = self.ensure_data()
        year_filter = str(ano or "Todos").strip()
        year_txt = "" if year_filter.lower() in ("todos", "todas", "all", "") else year_filter
        today = date.today()
        today_iso = today.isoformat()

        def _matches_year(enc: dict[str, Any]) -> bool:
            if not year_txt:
                return True
            candidates = [
                str(enc.get("data_entrega", "") or "").strip(),
                str(enc.get("data_criacao", "") or "").strip()[:10],
            ]
            for value in candidates:
                if len(value) >= 4 and value[:4] == year_txt:
                    return True
            return False

        def _safe_date_txt(value: Any) -> date | None:
            txt = str(value or "").strip()
            if not txt:
                return None
            try:
                return datetime.fromisoformat(txt[:10]).date()
            except Exception:
                return None

        def _phase_meta(
            *,
            enc_state: str,
            has_montagem: bool,
            montagem_estado: str,
            laser_status: str,
            shipping_status: str,
            has_guide: bool,
            trip_number: str,
            trip_state: str,
            transport_pending: bool,
            delivered: bool,
        ) -> tuple[str, str]:
            enc_norm = self.desktop_main.norm_text(enc_state)
            laser_norm = self.desktop_main.norm_text(laser_status)
            shipping_norm = self.desktop_main.norm_text(shipping_status)
            trip_norm = self.desktop_main.norm_text(trip_state)
            montagem_norm = self.desktop_main.norm_text(montagem_estado)
            if delivered or "entreg" in trip_norm:
                return "Entregue", "success"
            if trip_number:
                return "Em transporte", "info"
            if transport_pending:
                return "A aguardar transporte", "warning"
            if "totalmente expedida" in shipping_norm or has_guide:
                return "Expedição", "info"
            if has_montagem and "pendente" in montagem_norm and "concluido" in laser_norm:
                return "Montagem", "warning"
            if "concluido" in laser_norm:
                return "Pronta para expedição", "success"
            if "completo" in laser_norm:
                return "Laser planeado", "info"
            if "parcial" in laser_norm:
                return "Laser parcial", "warning"
            if "planear" in laser_norm:
                return "Preparação", "default"
            if "montag" in enc_norm:
                return "Montagem", "warning"
            if "produc" in enc_norm or "curso" in enc_norm:
                return "Em produção", "info"
            if "concl" in enc_norm:
                return "Concluída", "success"
            return "Preparação", "default"

        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(data.get("clientes", []) or [])
            if isinstance(row, dict)
        }
        deadline_rows = {
            str(row.get("numero", "") or "").strip(): dict(row)
            for row in list(self.planning_laser_deadline_rows() or [])
            if isinstance(row, dict)
        }
        delay_payload = self.pulse_plan_delay_rows(period_days=60, year_filter=year_txt or None, encomenda="Todas")
        delay_open = {
            str(item.get("numero", "") or "").strip(): dict(item)
            for item in list(delay_payload.get("items", []) or [])
            if isinstance(item, dict) and not bool(item.get("acknowledged"))
        }
        pending_transport_rows = {
            str(row.get("numero", "") or "").strip(): dict(row)
            for row in list(self.transport_pending_orders("") or [])
            if isinstance(row, dict)
        }
        active_trips: dict[str, dict[str, str]] = {}
        for trip in list(data.get("transportes", []) or []):
            if not isinstance(trip, dict):
                continue
            trip_num = str(trip.get("numero", "") or "").strip()
            trip_state = str(trip.get("estado", "") or "Planeado").strip() or "Planeado"
            if not trip_num or "anulad" in self.desktop_main.norm_text(trip_state):
                continue
            for stop in list(trip.get("paragens", []) or []):
                if not isinstance(stop, dict):
                    continue
                enc_num = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
                if not enc_num:
                    continue
                active_trips[enc_num] = {
                    "numero": trip_num,
                    "estado": self._transport_stop_state(stop, trip_state),
                }

        rows: list[dict[str, Any]] = []
        action_rows: list[dict[str, Any]] = []
        logistics_rows: list[dict[str, Any]] = []
        phase_counts: dict[str, int] = {}

        for enc in list(data.get("encomendas", []) or []):
            if not isinstance(enc, dict) or not _matches_year(enc):
                continue
            numero = str(enc.get("numero", "") or "").strip()
            if not numero:
                continue
            try:
                self.desktop_main.update_estado_expedicao_encomenda(enc)
            except Exception:
                pass

            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = clients.get(client_code, "") or str(enc.get("cliente_nome", "") or "").strip()
            client_display = " - ".join(part for part in [client_code, client_name] if part).strip(" -") or client_code or "-"

            deadline = dict(deadline_rows.get(numero, {}) or {})
            delay = dict(delay_open.get(numero, {}) or {})
            latest_guide = self._transport_latest_guide_for_order(numero) or {}
            active_trip = dict(active_trips.get(numero, {}) or {})
            trip_number = str(active_trip.get("numero", "") or str(enc.get("transporte_numero", "") or "")).strip()
            trip_state = str(active_trip.get("estado", "") or str(enc.get("estado_transporte", "") or "")).strip()
            has_guide = bool(str(latest_guide.get("numero", "") or "").strip())
            shipping_status = str(enc.get("estado_expedicao", "Não expedida") or "Não expedida").strip()
            transport_pending = numero in pending_transport_rows
            delivery_date = _safe_date_txt(enc.get("data_entrega", ""))
            delivery_txt = str(enc.get("data_entrega", "") or "").strip() or "-"
            delivery_overdue = bool(delivery_date and delivery_date < today and "entreg" not in self.desktop_main.norm_text(trip_state))

            montagem_items = list(self.desktop_main.encomenda_montagem_itens(enc) or [])
            has_montagem = bool(montagem_items)
            montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "Não aplicável").strip()
            montagem_shortages = list(self._order_montagem_shortages(enc) or []) if has_montagem else []
            shortage_count = len(montagem_shortages)

            laser_status = str(deadline.get("estado", "") or ("Sem laser" if not deadline else "-")).strip() or "-"
            laser_plan_txt = str(deadline.get("planeado_txt", "") or "-").strip() or "-"
            laser_end_txt = str(deadline.get("fim_txt", "") or "-").strip() or "-"
            phase_label, phase_tone = _phase_meta(
                enc_state=str(enc.get("estado", "") or "").strip(),
                has_montagem=has_montagem,
                montagem_estado=montagem_estado,
                laser_status=laser_status,
                shipping_status=shipping_status,
                has_guide=has_guide,
                trip_number=trip_number,
                trip_state=trip_state,
                transport_pending=transport_pending,
                delivered=bool("entreg" in self.desktop_main.norm_text(trip_state)),
            )
            phase_counts[phase_label] = phase_counts.get(phase_label, 0) + 1

            signal = "OK"
            signal_tone = "success" if phase_tone == "success" else "default"
            next_action = "-"
            if delivery_overdue:
                signal = "Entrega ultrapassada"
                signal_tone = "danger"
                next_action = "Rever prioridade real e contactar cliente se necessário."
            elif delay:
                signal = "Fora do planeamento"
                signal_tone = "danger"
                next_action = "Rever o plano do laser ou justificar o atraso."
            elif shortage_count > 0:
                signal = f"Falta montagem ({shortage_count})"
                signal_tone = "danger"
                next_action = "Validar stock e gerar reposição de montagem."
            elif transport_pending:
                signal = "Sem transporte"
                signal_tone = "warning"
                next_action = "Requisitar transporte ou subcontrato."
            elif "planear" in self.desktop_main.norm_text(laser_status):
                signal = "Por planear"
                signal_tone = "warning"
                next_action = "Planear o corte laser."
            elif "parcial" in self.desktop_main.norm_text(laser_status):
                signal = "Planeamento parcial"
                signal_tone = "warning"
                next_action = "Completar o planeamento do laser."
            elif has_montagem and "pendente" in self.desktop_main.norm_text(montagem_estado):
                signal = "Montagem pendente"
                signal_tone = "info"
                next_action = "Preparar consumos e fechar montagem."
            elif has_guide and not trip_number:
                signal = "Guia emitida"
                signal_tone = "info"
                next_action = "Confirmar saída / transporte."
            elif trip_number:
                signal = trip_state or "Em transporte"
                signal_tone = "info"
                next_action = "Acompanhar entrega."
            elif phase_label == "Pronta para expedição":
                signal = "Pronta a expedir"
                signal_tone = "success"
                next_action = "Emitir guia ou carregar transporte."

            row = {
                "numero": numero,
                "cliente": client_display,
                "estado": str(enc.get("estado", "") or "").strip() or "-",
                "fase": phase_label,
                "fase_tone": phase_tone,
                "laser": laser_status,
                "laser_planeado": laser_plan_txt,
                "laser_fim": laser_end_txt,
                "montagem": montagem_estado if has_montagem else "Não aplicável",
                "expedicao": shipping_status or "-",
                "guia_numero": str(latest_guide.get("numero", "") or "").strip(),
                "transporte_numero": trip_number,
                "transporte_estado": trip_state or "-",
                "transportadora": str(enc.get("transportadora_nome", "") or pending_transport_rows.get(numero, {}).get("transportadora_nome", "") or "-").strip() or "-",
                "zona": str(enc.get("zona_transporte", "") or pending_transport_rows.get(numero, {}).get("zona_transporte", "") or "-").strip() or "-",
                "peso_bruto_kg": round(self._parse_float(enc.get("peso_bruto_kg", pending_transport_rows.get(numero, {}).get("peso_bruto_kg", 0)), 0), 2),
                "paletes": round(self._parse_float(enc.get("paletes", pending_transport_rows.get(numero, {}).get("paletes", 0)), 0), 2),
                "entrega": delivery_txt,
                "sinal": signal,
                "signal_tone": signal_tone,
                "next_action": next_action,
                "delay_open": bool(delay),
                "delivery_overdue": delivery_overdue,
            }
            rows.append(row)
            if signal != "OK" or next_action != "-":
                action_rows.append(
                    {
                        "numero": numero,
                        "cliente": client_display,
                        "motivo": signal,
                        "acao": next_action,
                        "entrega": delivery_txt,
                        "tone": signal_tone,
                    }
                )
            if has_guide or trip_number or transport_pending or self._transport_is_own_cargo(enc):
                logistics_rows.append(
                    {
                        "numero": numero,
                        "cliente": client_display,
                        "guia": str(latest_guide.get("numero", "") or "-").strip() or "-",
                        "transporte": trip_number or "-",
                        "transportadora": str(row.get("transportadora", "-") or "-"),
                        "zona": str(row.get("zona", "-") or "-"),
                        "peso": f"{float(row.get('peso_bruto_kg', 0) or 0):.1f} kg",
                        "estado": signal if trip_number or transport_pending or has_guide else shipping_status or "-",
                        "tone": signal_tone if signal_tone != "default" else phase_tone,
                    }
                )

        phase_priority = {
            "Preparação": 0,
            "Laser parcial": 1,
            "Laser planeado": 2,
            "Montagem": 3,
            "Pronta para expedição": 4,
            "Expedição": 5,
            "A aguardar transporte": 6,
            "Em transporte": 7,
            "Entregue": 8,
            "Concluída": 9,
        }
        tone_priority = {"danger": 0, "warning": 1, "info": 2, "success": 3, "default": 4}
        rows.sort(
            key=lambda row: (
                tone_priority.get(str(row.get("signal_tone", "default")), 4),
                0 if bool(row.get("delivery_overdue")) else 1,
                str(row.get("entrega", "") or "9999-99-99"),
                phase_priority.get(str(row.get("fase", "") or ""), 99),
                str(row.get("numero", "") or ""),
            )
        )
        action_rows.sort(
            key=lambda row: (
                tone_priority.get(str(row.get("tone", "default")), 4),
                str(row.get("entrega", "") or "9999-99-99"),
                str(row.get("numero", "") or ""),
            )
        )
        logistics_rows.sort(
            key=lambda row: (
                tone_priority.get(str(row.get("tone", "default")), 4),
                str(row.get("numero", "") or ""),
            )
        )

        open_orders = sum(1 for row in rows if str(row.get("fase", "") or "") not in {"Entregue", "Concluída"})
        risk_count = sum(1 for row in rows if str(row.get("signal_tone", "") or "") == "danger")
        ready_shipping = sum(1 for row in rows if str(row.get("fase", "") or "") in {"Pronta para expedição", "Expedição"})
        waiting_transport = sum(1 for row in rows if str(row.get("fase", "") or "") == "A aguardar transporte")
        in_transport = sum(1 for row in rows if str(row.get("fase", "") or "") == "Em transporte")
        montagem_pending = sum(1 for row in rows if "Montagem" in str(row.get("fase", "") or ""))

        cards = [
            {"title": "Encomendas abertas", "value": str(open_orders), "subtitle": f"{len(rows)} no horizonte atual", "tone": "info"},
            {"title": "Em risco", "value": str(risk_count), "subtitle": "Atraso ao plano, falta ou entrega ultrapassada", "tone": "danger" if risk_count else "success"},
            {"title": "Prontas a expedir", "value": str(ready_shipping), "subtitle": "Laser e/ou montagem já resolvidos", "tone": "success" if ready_shipping else "default"},
            {"title": "A aguardar transporte", "value": str(waiting_transport), "subtitle": "Carga nossa sem viagem ativa", "tone": "warning" if waiting_transport else "default"},
            {"title": "Em transporte", "value": str(in_transport), "subtitle": "Viagens em curso", "tone": "info" if in_transport else "default"},
            {"title": "Montagem pendente", "value": str(montagem_pending), "subtitle": "Itens ainda por fechar em montagem", "tone": "warning" if montagem_pending else "default"},
        ]
        phase_rows = [
            {"fase": key, "total": value}
            for key, value in sorted(
                phase_counts.items(),
                key=lambda item: (phase_priority.get(str(item[0] or ""), 99), str(item[0] or "")),
            )
        ]
        return {
            "cards": cards,
            "phase_rows": phase_rows,
            "order_rows": rows[:24],
            "action_rows": action_rows[:14],
            "logistics_rows": logistics_rows[:14],
            "updated_at": str(self.desktop_main.now_iso() or "").strip(),
            "selected_year": year_filter or "Todos",
        }

    def dashboard_counts(self) -> list[dict[str, str]]:
        data = self.ensure_data()
        encomendas_abertas = sum(1 for enc in data.get("encomendas", []) if "concl" not in str(enc.get("estado", "")).lower())
        encomendas_montagem = sum(1 for enc in data.get("encomendas", []) if "montag" in self.desktop_main.norm_text(enc.get("estado", "")))
        material_disponivel = sum(
            max(0.0, self._parse_float(m.get("quantidade", 0), 0) - self._parse_float(m.get("reservado", 0), 0))
            for m in data.get("materiais", [])
        )
        return [
            {"title": "Materias", "value": str(len(data.get("materiais", []))), "subtitle": f"Disponivel {self._fmt(material_disponivel)}"},
            {"title": "Encomendas", "value": str(len(data.get("encomendas", []))), "subtitle": f"Abertas {encomendas_abertas} | Montagem {encomendas_montagem}"},
            {"title": "Clientes", "value": str(len(data.get("clientes", []))), "subtitle": "Base ativa"},
            {"title": "Fornecedores", "value": str(len(data.get("fornecedores", []))), "subtitle": "Compras e stock"},
        ]
