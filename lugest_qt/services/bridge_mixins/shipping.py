from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any


class ShippingBridgeMixin:
    """Shipping guide operations for the Qt bridge."""

    def expedicao_pending_orders(self, filter_text: str = "", estado: str = "Todas") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        estado_filter = str(estado or "Todas").strip()
        rows = []
        for enc in data.get("encomendas", []):
            self.desktop_main.update_estado_expedicao_encomenda(enc)
            pieces = list(self.desktop_main.encomenda_pecas(enc))
            disponivel = sum(max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(p), 0)) for p in pieces)
            if disponivel <= 0:
                continue
            estado_exp = str(enc.get("estado_expedicao", "Nao expedida") or "Nao expedida").strip()
            if estado_filter != "Todas" and estado_exp != estado_filter:
                continue
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_obj = {}
            find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
            if callable(find_cliente_fn):
                cli_obj = find_cliente_fn(data, cli_code) or {}
            cliente_txt = " - ".join([part for part in [cli_code, str(cli_obj.get("nome", "") or "").strip()] if part]).strip()
            row = {
                "numero": str(enc.get("numero", "") or "").strip(),
                "cliente": cliente_txt or cli_code or "-",
                "estado": str(enc.get("estado", "") or "").strip(),
                "estado_expedicao": estado_exp,
                "disponivel": round(disponivel, 1),
                "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: ((item.get("data_entrega") or "9999-99-99"), item.get("numero") or ""))
        return rows

    def expedicao_available_pieces(self, enc_num: str) -> list[dict[str, Any]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            return []
        rows = []
        for piece in self.desktop_main.encomenda_pecas(enc):
            self.desktop_main.ensure_peca_operacoes(piece)
            pronta = max(0.0, self._parse_float(getattr(self.desktop_main, "peca_qtd_pronta_expedicao")(piece), 0))
            disponivel = max(0.0, self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(piece), 0))
            if disponivel <= 0:
                continue
            rows.append(
                {
                    "id": str(piece.get("id", "") or "").strip(),
                    "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                    "descricao": str(
                        piece.get("Observacoes")
                        or piece.get("Observações")
                        or piece.get("descricao")
                        or piece.get("ref_externa")
                        or piece.get("ref_interna")
                        or ""
                    ).strip(),
                    "estado": str(piece.get("estado", "") or "").strip(),
                    "pronta_expedicao": self._fmt(pronta),
                    "qtd_expedida": self._fmt(piece.get("qtd_expedida", 0)),
                    "disponivel": self._fmt(disponivel),
                    "pronta_expedicao_num": pronta,
                    "qtd_expedida_num": self._parse_float(piece.get("qtd_expedida", 0), 0),
                    "disponivel_num": disponivel,
                    "material": str(piece.get("material", "") or "").strip(),
                    "espessura": str(piece.get("espessura", "") or "").strip(),
                    "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                }
            )
        rows.sort(key=lambda item: (item.get("ref_interna") or "", item.get("ref_externa") or ""))
        return rows

    def expedicao_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for ex in sorted(list(self.ensure_data().get("expedicoes", []) or []), key=lambda row: str(row.get("data_emissao", "") or ""), reverse=True):
            row = {
                "numero": str(ex.get("numero", "") or "").strip(),
                "tipo": str(ex.get("tipo", "") or "").strip(),
                "encomenda": str(ex.get("encomenda", "") or "").strip(),
                "cliente": str(ex.get("cliente_nome", "") or ex.get("cliente", "") or "").strip(),
                "data_emissao": str(ex.get("data_emissao", "") or "").replace("T", " ")[:19],
                "estado": "Anulada" if bool(ex.get("anulada")) else str(ex.get("estado", "") or "").strip(),
                "linhas": len(list(ex.get("linhas", []) or [])),
                "anulada": bool(ex.get("anulada")),
                "transportador": str(ex.get("transportador", "") or "").strip(),
                "matricula": str(ex.get("matricula", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        return rows

    def expedicao_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        lines = []
        for line in list(ex.get("linhas", []) or []):
            lines.append(
                {
                    "peca_id": str(line.get("peca_id", "") or "").strip(),
                    "ref_interna": str(line.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "qtd": self._fmt(line.get("qtd", 0)),
                    "peso": self._fmt(line.get("peso", 0)),
                    "manual": bool(line.get("manual")),
                    "encomenda": str(line.get("encomenda", "") or ex.get("encomenda", "") or "").strip(),
                }
            )
        return {
            "numero": str(ex.get("numero", "") or "").strip(),
            "tipo": str(ex.get("tipo", "") or "").strip(),
            "encomenda": str(ex.get("encomenda", "") or "").strip(),
            "cliente": str(ex.get("cliente_nome", "") or ex.get("cliente", "") or "").strip(),
            "estado": "Anulada" if bool(ex.get("anulada")) else str(ex.get("estado", "") or "").strip(),
            "data_emissao": str(ex.get("data_emissao", "") or "").replace("T", " ")[:19],
            "data_transporte": str(ex.get("data_transporte", "") or "").replace("T", " ")[:19],
            "transportador": str(ex.get("transportador", "") or "").strip(),
            "matricula": str(ex.get("matricula", "") or "").strip(),
            "destinatario": str(ex.get("destinatario", "") or "").strip(),
            "local_descarga": str(ex.get("local_descarga", "") or "").strip(),
            "observacoes": str(ex.get("observacoes", "") or "").strip(),
            "anulada_motivo": str(ex.get("anulada_motivo", "") or "").strip(),
            "lines": lines,
        }

    def expedicao_open_pdf(self, numero: str) -> Path:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        path = Path(tempfile.gettempdir()) / f"lugest_guia_{numero}.pdf"
        self.ne_expedicao_actions.render_expedicao_pdf(self, str(path), ex)
        os.startfile(str(path))
        return path

    def expedicao_render_pdf(self, numero: str, path: str | Path, include_all_vias: bool = False) -> Path:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        out_path = Path(path)
        self.ne_expedicao_actions.render_expedicao_pdf(self, str(out_path), ex, include_all_vias=include_all_vias)
        return out_path

    def _exp_validation_code(self, issue_date: str | None = None) -> str:
        issue_date = str(issue_date or self.desktop_main.now_iso())
        serie_guess = ""
        default_serie_fn = getattr(self.desktop_main, "_exp_default_serie_id", None)
        if callable(default_serie_fn):
            try:
                serie_guess = str(default_serie_fn("GT", issue_date) or "").strip()
            except Exception:
                serie_guess = ""
        find_series_fn = getattr(self.desktop_main, "_find_at_series", None)
        if not callable(find_series_fn):
            return ""
        try:
            serie_obj = find_series_fn(self.ensure_data(), doc_type="GT", serie_id=serie_guess) or {}
        except Exception:
            serie_obj = {}
        return str(serie_obj.get("validation_code", "") or "").strip()

    def expedicao_defaults_for_order(self, enc_num: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        cli_code = str(enc.get("cliente", "") or "").strip()
        cli = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn):
            cli = find_cliente_fn(self.ensure_data(), cli_code) or {}
        cli_nome = str(cli.get("nome", "") or cli_code or "").strip()
        emit_cfg = dict(self.desktop_main.get_guia_emitente_info() or {})
        rodape = list(self.desktop_main.get_empresa_rodape_lines() or [])
        local_carga = str(
            emit_cfg.get("local_carga", "")
            or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))
            or ""
        ).strip()
        return {
            "codigo_at": self._exp_validation_code(self.desktop_main.now_iso()),
            "tipo_via": "Original",
            "emitente_nome": str(emit_cfg.get("nome", "") or "").strip(),
            "emitente_nif": str(emit_cfg.get("nif", "") or "").strip(),
            "emitente_morada": str(emit_cfg.get("morada", "") or "").strip(),
            "destinatario": cli_nome,
            "dest_nif": str(cli.get("nif", "") or "").strip(),
            "dest_morada": str(cli.get("morada", "") or "").strip(),
            "local_carga": local_carga,
            "local_descarga": str(cli.get("morada", "") or "").strip(),
            "data_transporte": str(self.desktop_main.now_iso()),
            "transportador": "",
            "matricula": "",
            "observacoes": f"Expedicao da encomenda {enc_num}",
        }

    def expedicao_manual_defaults(self) -> dict[str, Any]:
        emit_cfg = dict(self.desktop_main.get_guia_emitente_info() or {})
        rodape = list(self.desktop_main.get_empresa_rodape_lines() or [])
        local_carga = str(
            emit_cfg.get("local_carga", "")
            or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))
            or ""
        ).strip()
        return {
            "codigo_at": self._exp_validation_code(self.desktop_main.now_iso()),
            "tipo_via": "Original",
            "emitente_nome": str(emit_cfg.get("nome", "") or "").strip(),
            "emitente_nif": str(emit_cfg.get("nif", "") or "").strip(),
            "emitente_morada": str(emit_cfg.get("morada", "") or "").strip(),
            "destinatario": "",
            "dest_nif": "",
            "dest_morada": "",
            "local_carga": local_carga,
            "local_descarga": "",
            "data_transporte": str(self.desktop_main.now_iso()),
            "transportador": "",
            "matricula": "",
            "observacoes": "",
        }

    def expedicao_product_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for prod in list(self.ensure_data().get("produtos", []) or []):
            qty = self._parse_float(prod.get("qty", 0), 0)
            row = {
                "codigo": str(prod.get("codigo", "") or "").strip(),
                "descricao": str(prod.get("descricao", "") or "").strip(),
                "qty": qty,
                "unid": str(prod.get("unid", "UN") or "UN").strip() or "UN",
                "qty_fmt": self._fmt(qty),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo") or "", item.get("descricao") or ""))
        return rows

    def expedicao_emit_off(self, enc_num: str, draft_lines: list[dict[str, Any]], guide_data: dict[str, Any]) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda n?o encontrada.")
        lines = [dict(line) for line in list(draft_lines or []) if isinstance(line, dict)]
        if not lines:
            raise ValueError("Sem linhas na guia.")
        requested_by_piece: dict[str, float] = {}
        pieces = {str(piece.get("id", "") or ""): piece for piece in self.desktop_main.encomenda_pecas(enc)}
        for line in lines:
            piece_id = str(line.get("peca_id", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            if not piece_id or qty <= 0:
                raise ValueError("Linha de guia invalida.")
            piece = pieces.get(piece_id)
            if piece is None:
                raise ValueError(f"Pe?a n?o encontrada para expedi??o: {piece_id}")
            requested_by_piece[piece_id] = requested_by_piece.get(piece_id, 0.0) + qty
        for piece_id, qty in requested_by_piece.items():
            available = self._parse_float(self.desktop_main.peca_qtd_disponivel_expedicao(pieces[piece_id]), 0)
            if qty > available + 1e-9:
                raise ValueError(f"Quantidade superior ao disponivel na peca {pieces[piece_id].get('ref_interna', piece_id)}.")
        cli_code = str(enc.get("cliente", "") or "").strip()
        cli = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn):
            cli = find_cliente_fn(self.ensure_data(), cli_code) or {}
        cli_nome = str(guide_data.get("destinatario", "") or cli.get("nome", "") or cli_code).strip()
        exp_ids, exp_err = self.desktop_main.next_expedicao_identifiers(
            self.ensure_data(),
            issue_date=str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            doc_type="GT",
            validation_code_hint=str(guide_data.get("codigo_at", "") or "").strip(),
        )
        if not exp_ids:
            raise ValueError(exp_err or "Nao foi possivel obter serie/ATCUD da guia.")
        ex_num = str(exp_ids.get("numero", "") or "").strip()
        ex = {
            "numero": ex_num,
            "tipo": "OFF",
            "encomenda": str(enc.get("numero", "") or "").strip(),
            "cliente": cli_code,
            "cliente_nome": cli_nome,
            "codigo_at": exp_ids.get("validation_code", ""),
            "serie_id": exp_ids.get("serie_id", ""),
            "seq_num": exp_ids.get("seq_num", 0),
            "at_validation_code": exp_ids.get("validation_code", ""),
            "atcud": exp_ids.get("atcud", ""),
            "tipo_via": str(guide_data.get("tipo_via", "Original") or "Original"),
            "emitente_nome": str(guide_data.get("emitente_nome", "") or ""),
            "emitente_nif": str(guide_data.get("emitente_nif", "") or ""),
            "emitente_morada": str(guide_data.get("emitente_morada", "") or ""),
            "destinatario": cli_nome,
            "dest_nif": str(guide_data.get("dest_nif", "") or cli.get("nif", "") or ""),
            "dest_morada": str(guide_data.get("dest_morada", "") or cli.get("morada", "") or ""),
            "local_carga": str(guide_data.get("local_carga", "") or ""),
            "local_descarga": str(guide_data.get("local_descarga", "") or ""),
            "data_emissao": self.desktop_main.now_iso(),
            "data_transporte": str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            "matricula": str(guide_data.get("matricula", "") or ""),
            "transportador": str(guide_data.get("transportador", "") or ""),
            "estado": "Emitida",
            "observacoes": str(guide_data.get("observacoes", "") or ""),
            "created_by": str((self.user or {}).get("username", "") or ""),
            "anulada": False,
            "anulada_motivo": "",
            "linhas": [],
        }
        for line in lines:
            piece_id = str(line.get("peca_id", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            piece = pieces[piece_id]
            piece["qtd_expedida"] = self._parse_float(piece.get("qtd_expedida", 0), 0) + qty
            piece.setdefault("expedicoes", []).append(ex_num)
            ex["linhas"].append(
                {
                    "encomenda": str(enc.get("numero", "") or "").strip(),
                    "peca_id": piece_id,
                    "ref_interna": str(line.get("ref_interna", "") or piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(line.get("ref_externa", "") or piece.get("ref_externa", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "qtd": qty,
                    "unid": str(line.get("unid", "UN") or "UN").strip() or "UN",
                    "peso": self._parse_float(line.get("peso", 0), 0),
                    "manual": False,
                }
            )
            try:
                self.desktop_main.atualizar_estado_peca(piece)
            except Exception:
                pass
        self.ensure_data().setdefault("expedicoes", []).append(ex)
        self.desktop_main.update_estado_expedicao_encomenda(enc)
        self._save(force=True)
        return self.expedicao_detail(ex_num)

    def expedicao_emit_manual(self, guide_data: dict[str, Any], lines: list[dict[str, Any]]) -> dict[str, Any]:
        clean_lines = [dict(line) for line in list(lines or []) if isinstance(line, dict)]
        if not clean_lines:
            raise ValueError("Sem linhas na guia.")
        products = {str(prod.get("codigo", "") or "").strip(): prod for prod in self.ensure_data().get("produtos", [])}
        for line in clean_lines:
            code = str(line.get("produto_codigo", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            if qty <= 0:
                raise ValueError("Quantidade invalida numa linha da guia.")
            if code:
                prod = products.get(code)
                if prod is None:
                    raise ValueError(f"Produto nao encontrado: {code}")
                if qty > self._parse_float(prod.get("qty", 0), 0) + 1e-9:
                    raise ValueError(f"Stock insuficiente para {code}.")
        exp_ids, exp_err = self.desktop_main.next_expedicao_identifiers(
            self.ensure_data(),
            issue_date=str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            doc_type="GT",
            validation_code_hint=str(guide_data.get("codigo_at", "") or "").strip(),
        )
        if not exp_ids:
            raise ValueError(exp_err or "Nao foi possivel obter serie/ATCUD da guia.")
        ex_num = str(exp_ids.get("numero", "") or "").strip()
        ex = {
            "numero": ex_num,
            "tipo": "Manual",
            "encomenda": "",
            "cliente": "",
            "cliente_nome": str(guide_data.get("destinatario", "") or "").strip(),
            "codigo_at": exp_ids.get("validation_code", ""),
            "serie_id": exp_ids.get("serie_id", ""),
            "seq_num": exp_ids.get("seq_num", 0),
            "at_validation_code": exp_ids.get("validation_code", ""),
            "atcud": exp_ids.get("atcud", ""),
            "tipo_via": str(guide_data.get("tipo_via", "Original") or "Original"),
            "emitente_nome": str(guide_data.get("emitente_nome", "") or ""),
            "emitente_nif": str(guide_data.get("emitente_nif", "") or ""),
            "emitente_morada": str(guide_data.get("emitente_morada", "") or ""),
            "destinatario": str(guide_data.get("destinatario", "") or "").strip(),
            "dest_nif": str(guide_data.get("dest_nif", "") or "").strip(),
            "dest_morada": str(guide_data.get("dest_morada", "") or "").strip(),
            "local_carga": str(guide_data.get("local_carga", "") or ""),
            "local_descarga": str(guide_data.get("local_descarga", "") or ""),
            "data_emissao": self.desktop_main.now_iso(),
            "data_transporte": str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso()),
            "matricula": str(guide_data.get("matricula", "") or ""),
            "transportador": str(guide_data.get("transportador", "") or ""),
            "estado": "Emitida",
            "observacoes": str(guide_data.get("observacoes", "") or ""),
            "created_by": str((self.user or {}).get("username", "") or ""),
            "anulada": False,
            "anulada_motivo": "",
            "linhas": [],
        }
        for line in clean_lines:
            code = str(line.get("produto_codigo", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            if code:
                prod = products.get(code)
                if prod is not None:
                    prod["qty"] = max(0.0, self._parse_float(prod.get("qty", 0), 0) - qty)
                    prod["atualizado_em"] = self.desktop_main.now_iso()
            ex["linhas"].append(
                {
                    "encomenda": "",
                    "peca_id": "",
                    "ref_interna": code,
                    "ref_externa": "",
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "qtd": qty,
                    "unid": str(line.get("unid", "UN") or "UN").strip() or "UN",
                    "peso": 0.0,
                    "manual": True,
                }
            )
        self.ensure_data().setdefault("expedicoes", []).append(ex)
        self._save(force=True)
        return self.expedicao_detail(ex_num)

    def expedicao_update(self, numero: str, guide_data: dict[str, Any]) -> dict[str, Any]:
        numero = str(numero or "").strip()
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        if bool(ex.get("anulada")):
            raise ValueError("Nao e possivel editar uma guia anulada.")
        ex["codigo_at"] = str(guide_data.get("codigo_at", "") or "").strip()
        ex["at_validation_code"] = str(guide_data.get("codigo_at", "") or "").strip()
        seq_num = int(self._parse_float(ex.get("seq_num", 0), 0) or 0)
        if ex.get("at_validation_code") and seq_num > 0:
            ex["atcud"] = f"{str(ex.get('at_validation_code', '')).strip()}-{seq_num}"
        ex["tipo_via"] = "Original"
        ex["emitente_nome"] = str(guide_data.get("emitente_nome", "") or "").strip()
        ex["emitente_nif"] = str(guide_data.get("emitente_nif", "") or "").strip()
        ex["emitente_morada"] = str(guide_data.get("emitente_morada", "") or "").strip()
        ex["destinatario"] = str(guide_data.get("destinatario", "") or "").strip()
        ex["dest_nif"] = str(guide_data.get("dest_nif", "") or "").strip()
        ex["dest_morada"] = str(guide_data.get("dest_morada", "") or "").strip()
        ex["local_carga"] = str(guide_data.get("local_carga", "") or "").strip()
        ex["local_descarga"] = str(guide_data.get("local_descarga", "") or "").strip()
        ex["data_transporte"] = str(guide_data.get("data_transporte", "") or self.desktop_main.now_iso())
        ex["transportador"] = str(guide_data.get("transportador", "") or "").strip()
        ex["matricula"] = str(guide_data.get("matricula", "") or "").strip()
        ex["observacoes"] = str(guide_data.get("observacoes", "") or "").strip()
        self._save(force=True)
        return self.expedicao_detail(numero)

    def expedicao_cancel(self, numero: str, reason: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        motivo = str(reason or "").strip()
        if not motivo:
            raise ValueError("E obrigatorio indicar justificacao.")
        ex = next((row for row in self.ensure_data().get("expedicoes", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if ex is None:
            raise ValueError("Guia n?o encontrada.")
        if bool(ex.get("anulada")):
            return self.expedicao_detail(numero)
        for line in list(ex.get("linhas", []) or []):
            if bool(line.get("manual")):
                code = str(line.get("ref_interna", "") or "").strip()
                qty = self._parse_float(line.get("qtd", 0), 0)
                if code:
                    prod = next((row for row in self.ensure_data().get("produtos", []) if str(row.get("codigo", "") or "").strip() == code), None)
                    if prod is not None:
                        prod["qty"] = self._parse_float(prod.get("qty", 0), 0) + qty
                        prod["atualizado_em"] = self.desktop_main.now_iso()
                continue
            enc = self.get_encomenda_by_numero(str(line.get("encomenda", "") or "").strip())
            if enc is None:
                continue
            piece_id = str(line.get("peca_id", "") or "").strip()
            qty = self._parse_float(line.get("qtd", 0), 0)
            for piece in self.desktop_main.encomenda_pecas(enc):
                if str(piece.get("id", "") or "").strip() == piece_id:
                    piece["qtd_expedida"] = max(0.0, self._parse_float(piece.get("qtd_expedida", 0), 0) - qty)
                    try:
                        self.desktop_main.atualizar_estado_peca(piece)
                    except Exception:
                        pass
                    break
            self.desktop_main.update_estado_expedicao_encomenda(enc)
        ex["anulada"] = True
        ex["estado"] = "Anulada"
        ex["anulada_motivo"] = motivo
        self._save(force=True)
        return self.expedicao_detail(numero)

