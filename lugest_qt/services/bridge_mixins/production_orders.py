from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Any


class ProductionOrdersBackendMixin:
    """Legacy adapter for production orders; see BACKEND_GUIDE.md."""

    def _find_piece_by_opp(self, opp: str) -> tuple[dict[str, Any], dict[str, Any]]:
        target = str(opp or "").strip()
        if not target:
            raise ValueError("OPP obrigatoria.")
        data = self.ensure_data()
        changed = False
        for enc in list(data.get("encomendas", []) or []):
            for piece in self.desktop_main.encomenda_pecas(enc):
                if not str(piece.get("opp", "") or "").strip():
                    piece["opp"] = self._next_order_opp_codigo(enc)
                    changed = True
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self._order_of_code(enc, create=True)
                    changed = True
                if str(piece.get("opp", "") or "").strip() == target:
                    if changed:
                        self._save(force=True)
                    return enc, piece
        if changed:
            self._save(force=True)
        raise ValueError("OPP não encontrada.")

    def _opp_rows_base(self) -> list[dict[str, Any]]:
        data = self.ensure_data()
        cliente_nome = {
            str(c.get("codigo", "") or "").strip(): str(c.get("nome", "") or "").strip()
            for c in list(data.get("clientes", []) or [])
            if isinstance(c, dict)
        }
        rows: list[dict[str, Any]] = []
        changed = False
        for enc in list(data.get("encomendas", []) or []):
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_display = f"{cli_code} - {cliente_nome.get(cli_code, '')}".strip(" -")
            enc_year = ""
            try:
                enc_year = str(
                    self.desktop_main._enc_extract_year(
                        enc.get("data_criacao", ""),
                        enc.get("data_entrega", ""),
                        enc.get("numero", ""),
                        enc.get("ano"),
                    )
                    or ""
                ).strip()
            except Exception:
                enc_year = ""
            if not enc_year:
                raw_delivery = str(enc.get("data_entrega", "") or "").strip()
                if len(raw_delivery) >= 4 and raw_delivery[:4].isdigit():
                    enc_year = raw_delivery[:4]
                else:
                    enc_year = str(datetime.now().year)
            for piece in self.desktop_main.encomenda_pecas(enc):
                if not str(piece.get("opp", "") or "").strip():
                    piece["opp"] = self.desktop_main.next_opp_numero(data)
                    changed = True
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self.desktop_main.next_of_numero(data)
                    changed = True
                ops = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
                qty_plan = self._parse_float(piece.get("quantidade_pedida", 0), 0)
                qty_ok = self._parse_float(piece.get("produzido_ok", 0), 0)
                qty_nok = self._parse_float(piece.get("produzido_nok", 0), 0)
                qty_qual = self._parse_float(piece.get("produzido_qualidade", 0), 0)
                qty_prod = qty_ok + qty_nok + qty_qual
                qty_exp = self._parse_float(piece.get("qtd_expedida", 0), 0)
                progress = round((qty_prod / qty_plan) * 100.0, 1) if qty_plan > 0 else 0.0
                running_ops = [op for op in ops if "produ" in self.desktop_main.norm_text(op.get("estado", ""))]
                pending_ops = [op for op in ops if "concl" not in self.desktop_main.norm_text(op.get("estado", ""))]
                current_op = ""
                if running_ops:
                    current_op = self.desktop_main.normalize_operacao_nome(running_ops[0].get("nome", ""))
                elif pending_ops:
                    current_op = self.desktop_main.normalize_operacao_nome(pending_ops[0].get("nome", ""))
                elif ops:
                    current_op = self.desktop_main.normalize_operacao_nome(ops[-1].get("nome", ""))
                current_operator = ""
                if running_ops:
                    current_operator = str(running_ops[0].get("user", "") or "").strip()
                if not current_operator:
                    hist_rows = list(piece.get("hist", []) or [])
                    if hist_rows:
                        current_operator = str(hist_rows[-1].get("user", "") or "").strip()
                tempo_real = self._parse_float(piece.get("tempo_producao_min", 0), 0)
                if tempo_real <= 0 and piece.get("inicio_producao") and not piece.get("fim_producao"):
                    tempo_real = self._parse_float(
                        self.desktop_main.iso_diff_minutes(piece.get("inicio_producao"), self.desktop_main.now_iso()),
                        0,
                    )
                ops_total = len([op for op in ops if str(op.get("nome", "") or "").strip()])
                ops_done = len([op for op in ops if self.desktop_main.operacao_esta_concluida(piece, op)])
                expedicoes = [str(num or "").strip() for num in list(piece.get("expedicoes", []) or []) if str(num or "").strip()]
                rows.append(
                    {
                        "opp": str(piece.get("opp", "") or "").strip(),
                        "of": str(piece.get("of", "") or "").strip(),
                        "piece_id": str(piece.get("id", "") or "").strip(),
                        "encomenda": str(enc.get("numero", "") or "").strip(),
                        "cliente": cli_display or cli_code or "-",
                        "cliente_codigo": cli_code,
                        "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                        "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                        "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
                        "material": str(piece.get("material", "") or "").strip(),
                        "espessura": str(piece.get("espessura", "") or "").strip(),
                        "estado": str(piece.get("estado", "") or "").strip(),
                        "operacao_atual": current_op or "-",
                        "operador_atual": current_operator or "-",
                        "qtd_plan": qty_plan,
                        "qtd_prod": qty_prod,
                        "qtd_exp": qty_exp,
                        "progress": progress,
                        "tempo_real": round(tempo_real, 2),
                        "ops_total": ops_total,
                        "ops_done": ops_done,
                        "ops_pending": max(0, ops_total - ops_done),
                        "inicio": str(piece.get("inicio_producao", "") or "").strip(),
                        "fim": str(piece.get("fim_producao", "") or "").strip(),
                        "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                        "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
                        "ano": enc_year,
                        "expedicoes": expedicoes,
                    }
                )
        if changed:
            self._save(force=True)
        rows.sort(key=lambda item: (item.get("opp", ""), item.get("encomenda", ""), item.get("ref_interna", "")))
        return rows

    def opp_rows(
        self,
        filter_text: str = "",
        estado: str = "Ativas",
        ano: str = "Todos",
        operacao: str = "Todas",
        cliente: str = "Todos",
    ) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        estado_filter = str(estado or "Ativas").strip().lower()
        ano_filter = str(ano or "Todos").strip()
        operacao_filter = str(operacao or "Todas").strip().lower()
        cliente_filter = str(cliente or "Todos").strip()
        rows: list[dict[str, Any]] = []
        for row in self._opp_rows_base():
            estado_norm = self.desktop_main.norm_text(row.get("estado", ""))
            if ano_filter.lower() not in ("todos", "todas", "all", "") and str(row.get("ano", "") or "").strip() != ano_filter:
                continue
            if cliente_filter.lower() not in ("todos", "todas", "all", ""):
                cliente_codigo = str(row.get("cliente_codigo", "") or "").strip()
                if cliente_codigo != cliente_filter.split(" - ", 1)[0].strip():
                    continue
            if operacao_filter not in ("todas", "todos", "all", ""):
                if operacao_filter not in self.desktop_main.norm_text(row.get("operacao_atual", "")):
                    continue
            if estado_filter not in ("todos", "todas", "all", ""):
                if "ativ" in estado_filter and "concl" in estado_norm and float(row.get("qtd_exp", 0) or 0) >= float(row.get("qtd_prod", 0) or 0):
                    continue
                elif "ativ" in estado_filter and "concl" in estado_norm:
                    continue
                if "prepar" in estado_filter and "prepar" not in estado_norm:
                    continue
                if ("curso" in estado_filter or "produc" in estado_filter) and ("produ" not in estado_norm and "incomplet" not in estado_norm):
                    continue
                if "concl" in estado_filter and "concl" not in estado_norm:
                    continue
                if "exped" in estado_filter and float(row.get("qtd_exp", 0) or 0) <= 0:
                    continue
                if "avaria" in estado_filter and "avari" not in estado_norm:
                    continue
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        return rows

    def opp_operations(self) -> list[str]:
        ops: set[str] = set()
        for row in self._opp_rows_base():
            op = str(row.get("operacao_atual", "") or "").strip()
            if op and op != "-":
                ops.add(op)
        return sorted(ops)

    def opp_client_portfolio(
        self,
        cliente: str = "Todos",
        ano: str = "Todos",
        filter_text: str = "",
    ) -> dict[str, Any]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        year_filter = str(ano or "Todos").strip()
        client_filter = str(cliente or "Todos").strip()
        client_code_filter = client_filter.split(" - ", 1)[0].strip()
        if client_filter.lower() in {"", "todos", "todas", "all"}:
            client_code_filter = ""

        client_names = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(data.get("clientes", []) or [])
            if isinstance(row, dict)
        }
        opp_rows = list(self._opp_rows_base())
        opp_by_order: dict[str, list[dict[str, Any]]] = {}
        for row in opp_rows:
            order_number = str(row.get("encomenda", "") or "").strip()
            if order_number:
                opp_by_order.setdefault(order_number, []).append(dict(row))

        billing_by_order: dict[str, list[dict[str, Any]]] = {}
        for row in list(self.billing_rows("", "Todas", "Todos") or []):
            order_number = str(row.get("encomenda_numero", "") or "").strip()
            if order_number:
                billing_by_order.setdefault(order_number, []).append(dict(row))

        all_orders: list[dict[str, Any]] = []
        years: set[str] = set()
        for order in list(data.get("encomendas", []) or []):
            if not isinstance(order, dict):
                continue
            order_number = str(order.get("numero", "") or "").strip()
            if not order_number:
                continue
            client_code = str(order.get("cliente", "") or "").strip()
            client_name = client_names.get(client_code, "")
            client_label = f"{client_code} - {client_name}".strip(" -") or "Sem cliente"
            try:
                order_year = str(
                    self.desktop_main._enc_extract_year(
                        order.get("data_criacao", ""),
                        order.get("data_entrega", ""),
                        order_number,
                        order.get("ano"),
                    )
                    or ""
                ).strip()
            except Exception:
                order_year = ""
            if not order_year:
                created = str(order.get("data_criacao", "") or "").strip()
                order_year = created[:4] if len(created) >= 4 and created[:4].isdigit() else str(datetime.now().year)
            years.add(order_year)
            if year_filter.lower() not in {"", "todos", "todas", "all"} and order_year != year_filter:
                continue

            pieces = [dict(row) for row in opp_by_order.get(order_number, [])]
            qty_plan = round(sum(self._parse_float(row.get("qtd_plan", 0), 0) for row in pieces), 2)
            qty_prod = round(sum(self._parse_float(row.get("qtd_prod", 0), 0) for row in pieces), 2)
            qty_exp = round(sum(self._parse_float(row.get("qtd_exp", 0), 0) for row in pieces), 2)
            progress = round((qty_prod / qty_plan) * 100.0, 1) if qty_plan > 0 else 0.0

            billing_rows = billing_by_order.get(order_number, [])
            sold = round(sum(self._parse_float(row.get("vendido", 0), 0) for row in billing_rows), 2)
            invoiced = round(sum(self._parse_float(row.get("faturado", 0), 0) for row in billing_rows), 2)
            received = round(sum(self._parse_float(row.get("recebido", 0), 0) for row in billing_rows), 2)
            if sold <= 0:
                sold = round(self._parse_float(order.get("valor_adjudicado", 0), 0), 2)
                if sold <= 0:
                    try:
                        sold = round(self._parse_float(self._billing_order_source(order).get("total", 0), 0), 2)
                    except Exception:
                        sold = 0.0
            pending_invoice = round(max(0.0, sold - invoiced), 2)
            receivable = round(max(0.0, invoiced - received), 2)
            production_states = {str(row.get("estado", "") or "").strip() for row in pieces if str(row.get("estado", "") or "").strip()}
            production_state = str(order.get("estado", "") or "").strip() or (", ".join(sorted(production_states)) if production_states else "Preparacao")
            order_row = {
                "encomenda": order_number,
                "of": next(
                    (str(row.get("of", "") or "").strip() for row in pieces if str(row.get("of", "") or "").strip()),
                    self._order_of_code(order),
                ),
                "orcamento": str(order.get("numero_orcamento", "") or "").strip(),
                "cliente_codigo": client_code,
                "cliente_nome": client_name,
                "cliente": client_label,
                "ano": order_year,
                "estado": production_state,
                "data_criacao": str(order.get("data_criacao", "") or "").replace("T", " ")[:16],
                "data_entrega": str(order.get("data_entrega", "") or "").replace("T", " ")[:10],
                "opp_count": len(pieces),
                "qtd_plan": qty_plan,
                "qtd_prod": qty_prod,
                "qtd_exp": qty_exp,
                "progress": progress,
                "adjudicado": sold,
                "faturado": invoiced,
                "recebido": received,
                "por_faturar": pending_invoice,
                "saldo_receber": receivable,
                "pieces": pieces,
            }
            if query and not any(query in str(value).lower() for key, value in order_row.items() if key != "pieces"):
                if not any(query in str(value).lower() for piece in pieces for value in piece.values()):
                    continue
            all_orders.append(order_row)

        client_groups: dict[str, dict[str, Any]] = {}
        for row in all_orders:
            code = str(row.get("cliente_codigo", "") or "").strip()
            key = code or "SEM_CLIENTE"
            target = client_groups.setdefault(
                key,
                {
                    "cliente_codigo": code,
                    "cliente_nome": str(row.get("cliente_nome", "") or "").strip(),
                    "cliente": str(row.get("cliente", "") or "Sem cliente").strip(),
                    "encomendas": 0,
                    "opp": 0,
                    "adjudicado": 0.0,
                    "faturado": 0.0,
                    "recebido": 0.0,
                    "por_faturar": 0.0,
                    "saldo_receber": 0.0,
                },
            )
            target["encomendas"] += 1
            target["opp"] += int(row.get("opp_count", 0) or 0)
            for field in ("adjudicado", "faturado", "recebido", "por_faturar", "saldo_receber"):
                target[field] = round(self._parse_float(target.get(field, 0), 0) + self._parse_float(row.get(field, 0), 0), 2)

        clients = sorted(client_groups.values(), key=lambda row: str(row.get("cliente", "") or "").lower())
        selected_orders = [
            row
            for row in all_orders
            if not client_code_filter or str(row.get("cliente_codigo", "") or "").strip() == client_code_filter
        ]
        selected_orders.sort(
            key=lambda row: (str(row.get("data_criacao", "") or ""), str(row.get("encomenda", "") or "")),
            reverse=True,
        )
        totals = {
            "clientes": len({str(row.get("cliente_codigo", "") or "") for row in selected_orders}),
            "encomendas": len(selected_orders),
            "opp": sum(int(row.get("opp_count", 0) or 0) for row in selected_orders),
            "adjudicado": round(sum(self._parse_float(row.get("adjudicado", 0), 0) for row in selected_orders), 2),
            "faturado": round(sum(self._parse_float(row.get("faturado", 0), 0) for row in selected_orders), 2),
            "recebido": round(sum(self._parse_float(row.get("recebido", 0), 0) for row in selected_orders), 2),
            "por_faturar": round(sum(self._parse_float(row.get("por_faturar", 0), 0) for row in selected_orders), 2),
            "saldo_receber": round(sum(self._parse_float(row.get("saldo_receber", 0), 0) for row in selected_orders), 2),
        }
        return {
            "clients": clients,
            "orders": selected_orders,
            "totals": totals,
            "years": sorted(years, reverse=True),
            "selected_client": client_code_filter,
        }

    def opp_detail(self, opp: str) -> dict[str, Any]:
        enc, piece = self._find_piece_by_opp(opp)
        ops = list(self.desktop_main.ensure_peca_operacoes(piece) or [])
        qty_plan = self._parse_float(piece.get("quantidade_pedida", 0), 0)
        qty_ok = self._parse_float(piece.get("produzido_ok", 0), 0)
        qty_nok = self._parse_float(piece.get("produzido_nok", 0), 0)
        qty_qual = self._parse_float(piece.get("produzido_qualidade", 0), 0)
        qty_prod = qty_ok + qty_nok + qty_qual
        qty_exp = self._parse_float(piece.get("qtd_expedida", 0), 0)
        progress = round((qty_prod / qty_plan) * 100.0, 1) if qty_plan > 0 else 0.0
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_obj = {}
        if cliente_codigo:
            try:
                cliente_obj = self.desktop_main.find_cliente(self.ensure_data(), cliente_codigo) or {}
            except Exception:
                cliente_obj = {}
        op_rows: list[dict[str, Any]] = []
        for op in ops:
            nome = self.desktop_main.normalize_operacao_nome(op.get("nome", ""))
            capacidade = self.desktop_main.operacao_input_qtd(piece, nome) if nome else 0.0
            qtd_total = self.desktop_main.operacao_qtd_total(op, fallback_done=capacidade)
            op_progress = round((qtd_total / capacidade) * 100.0, 1) if capacidade > 0 else 0.0
            op_rows.append(
                {
                    "nome": nome or "-",
                    "estado": str(op.get("estado", "") or "").strip() or "Pendente",
                    "user": str(op.get("user", "") or "").strip(),
                    "inicio": str(op.get("inicio", "") or "").replace("T", " ")[:19],
                    "fim": str(op.get("fim", "") or "").replace("T", " ")[:19],
                    "qtd_ok": self._fmt(op.get("qtd_ok", 0)),
                    "qtd_nok": self._fmt(op.get("qtd_nok", 0)),
                    "qtd_qual": self._fmt(op.get("qtd_qual", 0)),
                    "capacidade": self._fmt(capacidade),
                    "progress": op_progress,
                }
            )
        event_rows: list[dict[str, Any]] = []
        target_piece_id = str(piece.get("id", "") or "").strip()
        target_ref = str(piece.get("ref_interna", "") or "").strip()
        target_enc = str(enc.get("numero", "") or "").strip()
        for ev in list(self.ensure_data().get("op_eventos", []) or []):
            if not isinstance(ev, dict):
                continue
            ev_piece = str(ev.get("peca_id", "") or "").strip()
            ev_ref = str(ev.get("ref_interna", "") or "").strip()
            ev_enc = str(ev.get("encomenda_numero", "") or "").strip()
            if target_piece_id and ev_piece == target_piece_id:
                pass
            elif target_ref and ev_ref == target_ref and ev_enc == target_enc:
                pass
            else:
                continue
            event_rows.append(
                {
                    "data": str(ev.get("created_at", "") or "").replace("T", " ")[:19],
                    "evento": str(ev.get("evento", "") or "").strip(),
                    "operacao": str(ev.get("operacao", "") or "").strip(),
                    "operador": str(ev.get("operador", "") or "").strip(),
                    "qtd_ok": self._fmt(ev.get("qtd_ok", 0)),
                    "qtd_nok": self._fmt(ev.get("qtd_nok", 0)),
                    "info": str(ev.get("info", "") or "").strip(),
                }
            )
        for ev in list(piece.get("hist", []) or []):
            if not isinstance(ev, dict):
                continue
            event_rows.append(
                {
                    "data": str(ev.get("ts", "") or "").replace("T", " ")[:19],
                    "evento": str(ev.get("acao", "") or "").strip(),
                    "operacao": " + ".join(str(item or "").strip() for item in list(ev.get("operacoes", []) or []) if str(item or "").strip()),
                    "operador": str(ev.get("user", "") or "").strip(),
                    "qtd_ok": self._fmt(ev.get("ok", 0)),
                    "qtd_nok": self._fmt(ev.get("nok", 0)),
                    "info": str(ev.get("motivo", "") or ev.get("inicio", "") or "").strip(),
                }
            )
        event_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        exp_rows: list[dict[str, Any]] = []
        for ex in list(self.ensure_data().get("expedicoes", []) or []):
            if not isinstance(ex, dict):
                continue
            for line in list(ex.get("linhas", []) or []):
                if str(line.get("peca_id", "") or "").strip() != target_piece_id:
                    continue
                exp_rows.append(
                    {
                        "guia": str(ex.get("numero", "") or "").strip(),
                        "data": str(ex.get("data_transporte", "") or ex.get("data_emissao", "") or "").replace("T", " ")[:19],
                        "estado": "Anulada" if bool(ex.get("anulada")) else str(ex.get("estado", "") or "").strip(),
                        "destinatario": str(ex.get("destinatario", "") or "").strip(),
                        "qtd": self._fmt(line.get("qtd", 0)),
                        "obs": str(ex.get("observacoes", "") or "").strip(),
                    }
                )
        exp_rows.sort(key=lambda item: str(item.get("data", "") or ""), reverse=True)
        return {
            "opp": str(piece.get("opp", "") or "").strip(),
            "of": str(piece.get("of", "") or "").strip(),
            "piece_id": target_piece_id,
            "encomenda": target_enc,
            "cliente": cliente_codigo,
            "cliente_nome": str(cliente_obj.get("nome", "") or "").strip(),
            "ref_interna": target_ref,
            "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
            "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
            "material": str(piece.get("material", "") or "").strip(),
            "espessura": str(piece.get("espessura", "") or "").strip(),
            "estado": str(piece.get("estado", "") or "").strip(),
            "operacoes": op_rows,
            "events": event_rows,
            "expedicoes": exp_rows,
            "qtd_plan": self._fmt(qty_plan),
            "qtd_prod": self._fmt(qty_prod),
            "qtd_exp": self._fmt(qty_exp),
            "qtd_ok": self._fmt(qty_ok),
            "qtd_nok": self._fmt(qty_nok),
            "qtd_qual": self._fmt(qty_qual),
            "progress": progress,
            "tempo_real": self._fmt(piece.get("tempo_producao_min", 0)),
            "inicio": str(piece.get("inicio_producao", "") or "").replace("T", " ")[:19],
            "fim": str(piece.get("fim_producao", "") or "").replace("T", " ")[:19],
            "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
            "expedida": float(qty_exp) > 0,
        }

    def opp_open_drawing(self, opp: str) -> str:
        enc, piece = self._find_piece_by_opp(opp)
        return self.operator_open_drawing(str(enc.get("numero", "") or "").strip(), str(piece.get("id", "") or "").strip())

    def opp_open_pdf(self, opp: str) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        enc, piece = self._find_piece_by_opp(opp)
        source_posto = self._operator_posto_for_operation(str(piece.get("operacao_atual", "") or "").strip()) or "Geral"
        row = self._operator_label_row(enc, piece, source_posto=source_posto)
        target = self._operator_label_tmp_path(str(enc.get("numero", "") or "").strip(), "opp_label")
        width, height = (150 * mm, 100 * mm)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(width, height))
        self._draw_operator_unit_label(canvas_obj, width, height, row, palette, logo_path, printed_at)
        canvas_obj.save()
        os.startfile(str(target))
        return target
