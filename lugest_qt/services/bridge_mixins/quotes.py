from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text


class QuotesBridgeMixin:
    """Quote, assembly model, and quote-to-order operations for the Qt bridge."""

    def _peek_next_orc_number(self) -> str:
        data = self.ensure_data()
        try:
            seq = int(data.get("orc_seq", 1) or 1)
        except Exception:
            seq = 1
        year = int(getattr(self.desktop_main.datetime.now(), "year", 0) or 0)
        return f"ORC-{year}-{seq:04d}"

    def _orc_number_sort_key(self, numero: str) -> tuple[int, int, str]:
        raw = str(numero or "").strip()
        parts = raw.split("-")
        year = 0
        seq = 0
        if len(parts) >= 3:
            try:
                year = int(parts[1])
            except Exception:
                year = 0
            try:
                seq = int(parts[2])
            except Exception:
                seq = 0
        return (year, seq, raw)

    def orc_next_number(self) -> str:
        return self._peek_next_orc_number()

    def _normalize_orc_client(self, value: Any) -> dict[str, str]:
        return dict(self.desktop_main._normalize_orc_cliente(value, self.ensure_data()) or {})

    def orc_rows(self, filter_text: str = "", state_filter: str = "Ativas", year: str = "Todos") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        year_raw = str(year or "Todos").strip()
        rows: list[dict[str, Any]] = []
        for raw in list(data.get("orcamentos", []) or []):
            if not isinstance(raw, dict):
                continue
            orc = dict(raw)
            client = self._normalize_orc_client(orc.get("cliente", {}))
            estado = str(orc.get("estado", "") or "").strip() or "Em edicao"
            estado_norm = self.desktop_main.norm_text(estado)
            row_year = str(self.orc_actions._orc_extract_year(orc.get("data", ""), orc.get("numero", ""), orc.get("ano")) or "").strip()
            if year_raw and year_raw.lower() not in {"todos", "todas", "all"} and row_year != year_raw:
                continue
            if state_raw and state_raw not in {"todos", "todas", "all"}:
                if "ativ" in state_raw and ("rejeitado" in estado_norm or "convertido" in estado_norm):
                    continue
                if "edi" in state_raw and "edi" not in estado_norm:
                    continue
                if "enviado" in state_raw and "enviado" not in estado_norm:
                    continue
                if "aprovado" in state_raw and "aprovado" not in estado_norm:
                    continue
                if "rejeitado" in state_raw and "rejeitado" not in estado_norm:
                    continue
                if "convertido" in state_raw and "convertido" not in estado_norm:
                    continue
            client_label = f"{client.get('codigo', '')} - {client.get('nome', '')}".strip(" -")
            row = {
                "numero": str(orc.get("numero", "") or "").strip(),
                "cliente": client_label or str(client.get("nome", "") or "").strip() or str(orc.get("cliente", "") or "").strip(),
                "estado": estado,
                "numero_encomenda": str(orc.get("numero_encomenda", "") or "").strip(),
                "total": round(self._parse_float(orc.get("total", 0), 0), 2),
                "data": str(orc.get("data", "") or "").strip()[:10],
                "linhas": len(list(orc.get("linhas", []) or [])),
                "ano": row_year,
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: self._orc_number_sort_key(str(item.get("numero", "") or "")), reverse=True)
        return rows

    def orc_available_years(self) -> list[str]:
        current_year = str(self.desktop_main.datetime.now().year)
        years = {current_year}
        for row in list(self.ensure_data().get("orcamentos", []) or []):
            if not isinstance(row, dict):
                continue
            year = str(self.orc_actions._orc_extract_year(row.get("data", ""), row.get("numero", ""), row.get("ano")) or "").strip()
            if year:
                years.add(year)
        return sorted(years, key=lambda value: int(value) if value.isdigit() else 0, reverse=True)

    def _find_orc_record(self, numero: str) -> dict[str, Any] | None:
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("orcamentos", []) or [])
                if str(row.get("numero", "") or "").strip() == numero_txt
            ),
            None,
        )

    def _json_safe_clone(self, payload: Any) -> Any:
        try:
            return json.loads(json.dumps(payload, ensure_ascii=False, default=str))
        except Exception:
            if isinstance(payload, dict):
                return {str(key): self._json_safe_clone(value) for key, value in payload.items()}
            if isinstance(payload, (list, tuple, set)):
                return [self._json_safe_clone(value) for value in payload]
            return payload

    def _ensure_orc_nesting_studies_table(self, conn: Any) -> None:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS orc_nesting_studies (
                    quote_number VARCHAR(80) NOT NULL,
                    group_key VARCHAR(190) NOT NULL,
                    group_label VARCHAR(255) NULL,
                    study_json LONGTEXT NULL,
                    created_at DATETIME NULL,
                    updated_at DATETIME NULL,
                    PRIMARY KEY (quote_number, group_key)
                )
                """
            )

    def _mysql_orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            return {}
        conn = None
        studies: dict[str, Any] = {}
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if not callable(connect):
                return {}
            conn = connect()
            self._ensure_orc_nesting_studies_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT group_key, study_json
                    FROM orc_nesting_studies
                    WHERE quote_number=%s
                    ORDER BY updated_at DESC, group_key ASC
                    """,
                    (numero_txt,),
                )
                rows = list(cur.fetchall() or [])
            for row in rows:
                group_key = str((row.get("group_key") if isinstance(row, dict) else row[0]) or "").strip()
                raw = row.get("study_json") if isinstance(row, dict) else row[1]
                if not group_key:
                    continue
                if isinstance(raw, (bytes, bytearray)):
                    raw = raw.decode("utf-8", errors="ignore")
                try:
                    parsed = json.loads(str(raw or "{}"))
                except Exception:
                    parsed = {}
                if isinstance(parsed, dict):
                    studies[group_key] = parsed
        except Exception:
            studies = {}
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
        return studies

    def _mysql_save_orc_nesting_study(self, numero: str, group_key: str, group_label: str, payload: dict[str, Any]) -> None:
        numero_txt = str(numero or "").strip()
        group_key_txt = str(group_key or "").strip()
        if not numero_txt or not group_key_txt:
            return
        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if not callable(connect):
                return
            conn = connect()
            self._ensure_orc_nesting_studies_table(conn)
            clean = json.dumps(self._json_safe_clone(payload), ensure_ascii=False)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO orc_nesting_studies (
                        quote_number,
                        group_key,
                        group_label,
                        study_json,
                        created_at,
                        updated_at
                    )
                    VALUES (%s, %s, %s, %s, NOW(), NOW())
                    ON DUPLICATE KEY UPDATE
                        group_label=VALUES(group_label),
                        study_json=VALUES(study_json),
                        updated_at=VALUES(updated_at)
                    """,
                    (numero_txt, group_key_txt, str(group_label or "").strip(), clean),
                )
            conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def _mysql_delete_orc_nesting_studies(self, numero: str, group_key: str = "") -> None:
        numero_txt = str(numero or "").strip()
        group_key_txt = str(group_key or "").strip()
        if not numero_txt:
            return
        conn = None
        try:
            connect = getattr(self.desktop_main, "_mysql_connect", None)
            if not callable(connect):
                return
            conn = connect()
            self._ensure_orc_nesting_studies_table(conn)
            with conn.cursor() as cur:
                if group_key_txt:
                    cur.execute(
                        "DELETE FROM orc_nesting_studies WHERE quote_number=%s AND group_key=%s",
                        (numero_txt, group_key_txt),
                    )
                else:
                    cur.execute("DELETE FROM orc_nesting_studies WHERE quote_number=%s", (numero_txt,))
            conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        orc = self._find_orc_record(numero_txt)
        local_studies = {}
        if isinstance(orc, dict):
            local_studies = {
                str(key): self._json_safe_clone(value)
                for key, value in dict(orc.get("nesting_studies", {}) or {}).items()
                if str(key).strip()
            }
        remote_studies = self._mysql_orc_nesting_studies(numero_txt)
        if not remote_studies:
            return local_studies
        merged = dict(local_studies)
        for group_key, remote_value in remote_studies.items():
            local_value = dict(merged.get(group_key, {}) or {})
            remote_updated = str(dict(remote_value or {}).get("updated_at", "") or "").strip()
            local_updated = str(local_value.get("updated_at", "") or "").strip()
            if not local_value or remote_updated >= local_updated:
                merged[group_key] = self._json_safe_clone(remote_value)
        return merged

    def orc_save_nesting_study(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        orc = self._find_orc_record(numero_txt)
        if orc is None:
            raise ValueError("Guarda primeiro o orçamento para associar o estudo de nesting.")
        clean = dict(self._json_safe_clone(payload) or {})
        group_key = str(clean.get("group_key", "") or "").strip()
        if not group_key:
            raise ValueError("Grupo de nesting inválido.")
        previous = dict(dict(orc.get("nesting_studies", {}) or {}).get(group_key, {}) or {})
        clean["quote_number"] = numero_txt
        clean["group_key"] = group_key
        clean["group_label"] = str(clean.get("group_label", previous.get("group_label", "")) or "").strip()
        clean["created_at"] = str(previous.get("created_at", "") or clean.get("created_at", "") or self.desktop_main.now_iso()).strip()
        clean["updated_at"] = self.desktop_main.now_iso()
        orc.setdefault("nesting_studies", {})[group_key] = clean
        orc["latest_nesting_bridge"] = dict(clean.get("quote_bridge", {}) or {})
        orc["latest_nesting_group_key"] = group_key
        orc["latest_nesting_updated_at"] = clean["updated_at"]
        self._save(force=True)
        try:
            self._mysql_save_orc_nesting_study(numero_txt, group_key, clean.get("group_label", ""), clean)
        except Exception:
            pass
        return self._json_safe_clone(clean)

    def orc_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        client = self._normalize_orc_client(orc.get("cliente", {}))
        lines: list[dict[str, Any]] = []
        for row in list(orc.get("linhas", []) or []):
            snapshot = self._quote_line_operation_snapshot(row, quote_number=numero, quote_state=str(orc.get("estado", "") or "").strip())
            raw_operacao = str(row.get("operacao", "") or "").strip()
            current_time = round(self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0), 4)
            current_price = round(self._parse_float(row.get("preco_unit", 0), 0), 4)
            derived_laser_base = (
                self.desktop_main.orc_line_is_piece(row)
                and bool(str(row.get("desenho", "") or "").strip())
                and "corte laser" in self.desktop_main.norm_text(raw_operacao)
                and (
                    current_time > 0
                    or current_price > 0
                )
            )
            laser_base_active = bool(row.get("laser_base_active", False) or derived_laser_base)
            laser_base_tempo = round(
                self._parse_float(
                    row.get(
                        "laser_base_tempo_unit",
                        current_time if laser_base_active else 0,
                    ),
                    0,
                ),
                4,
            )
            laser_base_preco = round(
                self._parse_float(
                    row.get(
                        "laser_base_preco_unit",
                        current_price if laser_base_active else 0,
                    ),
                    0,
                ),
                4,
            )
            display_extra_time_map = self._quote_collect_non_laser_map(dict(snapshot.get("tempos_operacao", {}) or {}), digits=4)
            display_extra_price_map = self._quote_collect_non_laser_map(dict(snapshot.get("custos_operacao", {}) or {}), digits=4)
            display_extra_time = round(sum(display_extra_time_map.values()), 4)
            display_extra_price = round(sum(display_extra_price_map.values()), 4)
            repair_extra_time_map = self._quote_collect_non_laser_map(
                dict(row.get("tempos_operacao", {}) or {}),
                dict(snapshot.get("tempos_operacao", {}) or {}),
                digits=4,
            )
            repair_extra_price_map = self._quote_collect_non_laser_map(
                dict(row.get("custos_operacao", {}) or {}),
                dict(snapshot.get("custos_operacao", {}) or {}),
                digits=4,
            )
            repair_extra_time = round(sum(repair_extra_time_map.values()), 4)
            repair_extra_price = round(sum(repair_extra_price_map.values()), 4)
            if laser_base_active:
                max_safe_base_time = round(max(0.0, current_time - repair_extra_time), 4)
                max_safe_base_price = round(max(0.0, current_price - repair_extra_price), 4)
                if laser_base_tempo > max_safe_base_time + 0.0001:
                    laser_base_tempo = max_safe_base_time
                if laser_base_preco > max_safe_base_price + 0.0001:
                    laser_base_preco = max_safe_base_price
            display_time = round(current_time, 2)
            display_price = round(current_price, 4)
            if laser_base_active:
                display_time = round(laser_base_tempo + display_extra_time, 2)
                display_price = round(laser_base_preco + display_extra_price, 4)
            line_qty = round(self._parse_float(row.get("qtd", 0), 0), 2)
            material_supplied_by_client = bool(row.get("material_supplied_by_client", False) or row.get("material_fornecido_cliente", False))
            lines.append(
                {
                    "tipo_item": self.desktop_main.normalize_orc_line_type(row.get("tipo_item")),
                    "ref_interna": str(row.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(row.get("ref_externa", "") or "").strip(),
                    "descricao": str(row.get("descricao", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "material_family": str(row.get("material_family", "") or "").strip(),
                    "material_subtype": str(row.get("material_subtype", "") or "").strip(),
                    "stock_item_kind": str(row.get("stock_item_kind", "") or "").strip(),
                    "material_supplied_by_client": material_supplied_by_client,
                    "material_fornecido_cliente": material_supplied_by_client,
                    "material_cost_included": (False if material_supplied_by_client else bool(row.get("material_cost_included", True))),
                    "espessura": self._fmt(row.get("espessura", "")),
                    "operacao": str(row.get("operacao", "") or "").strip(),
                    "produto_codigo": str(row.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(row.get("produto_unid", "") or "").strip(),
                    "_product_pending_create": bool(row.get("_product_pending_create", False)),
                    "conjunto_codigo": str(row.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(row.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(row.get("grupo_uuid", "") or "").strip(),
                    "qtd_base": round(self._parse_float(row.get("qtd_base", row.get("qtd", 0)), 0), 2),
                    "tempo_peca_min": display_time,
                    "qtd": line_qty,
                    "preco_unit": display_price,
                    "total": round(line_qty * display_price, 2),
                    "desenho": str(row.get("desenho", "") or "").strip(),
                    "laser_base_active": laser_base_active,
                    "laser_base_tempo_unit": laser_base_tempo,
                    "laser_base_preco_unit": laser_base_preco,
                    "operacoes_lista": list(snapshot.get("operacoes", []) or []),
                    "operacoes_fluxo": [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)],
                    "operacoes_detalhe": [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                    "tempos_operacao": dict(snapshot.get("tempos_operacao", {}) or {}),
                    "custos_operacao": dict(snapshot.get("custos_operacao", {}) or {}),
                    "quote_cost_snapshot": dict(snapshot.get("quote_cost_snapshot", {}) or {}),
                    "stock_material_id": str(row.get("stock_material_id", "") or "").strip(),
                    "price_per_kg": round(self._parse_float(row.get("price_per_kg", 0), 0), 4),
                    "price_base_value": round(self._parse_float(row.get("price_base_value", 0), 0), 4),
                    "price_base_label": str(row.get("price_base_label", "") or "").strip(),
                    "price_markup_pct": round(self._parse_float(row.get("price_markup_pct", 0), 0), 2),
                    "stock_metric_value": round(self._parse_float(row.get("stock_metric_value", 0), 0), 4),
                    "meters_per_unit": round(self._parse_float(row.get("meters_per_unit", 0), 0), 3),
                    "kg_per_m": round(self._parse_float(row.get("kg_per_m", 0), 0), 4),
                    "length_mm": round(self._parse_float(row.get("length_mm", 0), 0), 1),
                    "width_mm": round(self._parse_float(row.get("width_mm", 0), 0), 1),
                    "thickness_mm": round(self._parse_float(row.get("thickness_mm", 0), 0), 2),
                    "diameter_mm": round(self._parse_float(row.get("diameter_mm", 0), 0), 1),
                    "profile_section": str(row.get("profile_section", "") or "").strip(),
                    "profile_size": str(row.get("profile_size", "") or "").strip(),
                    "tube_section": str(row.get("tube_section", "") or "").strip(),
                    "quality": str(row.get("quality", "") or "").strip(),
                    "calc_mode": str(row.get("calc_mode", "") or "").strip(),
                }
            )
        return {
            "numero": str(orc.get("numero", "") or "").strip(),
            "data": str(orc.get("data", "") or "").strip()[:10],
            "estado": str(orc.get("estado", "") or "").strip() or "Em edicao",
            "cliente": client,
            "posto_trabalho": self._normalize_workcenter_value(orc.get("posto_trabalho", "")),
            "iva_perc": round(self._parse_float(orc.get("iva_perc", 23), 23), 2),
            "desconto_perc": round(self._parse_float(orc.get("desconto_perc", 0), 0), 2),
            "desconto_valor": round(self._parse_float(orc.get("desconto_valor", 0), 0), 2),
            "subtotal_bruto": round(self._parse_float(orc.get("subtotal_bruto", 0), 0), 2),
            "preco_transporte": round(self._parse_float(orc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(orc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self._parse_float(orc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(orc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(orc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(orc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(orc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(orc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": str(orc.get("zona_transporte", "") or "").strip(),
            "subtotal": round(self._parse_float(orc.get("subtotal", 0), 0), 2),
            "total": round(self._parse_float(orc.get("total", 0), 0), 2),
            "numero_encomenda": str(orc.get("numero_encomenda", "") or "").strip(),
            "executado_por": str(orc.get("executado_por", "") or "").strip(),
            "nota_transporte": str(orc.get("nota_transporte", "") or "").strip(),
            "notas_pdf": str(orc.get("notas_pdf", "") or "").strip(),
            "nota_cliente": str(orc.get("nota_cliente", "") or "").strip(),
            "nesting_bridge": dict(orc.get("latest_nesting_bridge", {}) or {}),
            "nesting_group_key": str(orc.get("latest_nesting_group_key", "") or "").strip(),
            "nesting_updated_at": str(orc.get("latest_nesting_updated_at", "") or "").strip(),
            "linhas": lines,
        }

    def orc_clients(self) -> list[dict[str, str]]:
        return list(self.order_clients())

    def _product_lookup(self, codigo: str) -> dict[str, Any] | None:
        code = str(codigo or "").strip()
        if not code:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("produtos", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )

    def _next_assembly_model_code(self) -> str:
        highest = 0
        for row in list(self.ensure_data().get("conjuntos_modelo", []) or []):
            codigo = str((row or {}).get("codigo", "") or "").strip().upper()
            digits = "".join(ch for ch in codigo if ch.isdigit())
            if digits:
                try:
                    highest = max(highest, int(digits))
                except Exception:
                    continue
        return f"CJ{highest + 1:04d}"

    def _normalize_assembly_model_item(self, payload: dict[str, Any]) -> dict[str, Any]:
        item_type = self.desktop_main.normalize_orc_line_type(payload.get("tipo_item"))
        quantity = round(self._parse_float(payload.get("qtd", 0), 0), 2)
        if quantity <= 0:
            raise ValueError("Quantidade invalida no conjunto.")
        stock_item_kind = str(payload.get("stock_item_kind", "") or "").strip()
        if item_type == self.desktop_main.ORC_LINE_TYPE_PIECE and (
            stock_item_kind == "raw_material" or str(payload.get("stock_material_id", "") or "").strip()
        ):
            stock_item_kind = "raw_material"
        elif item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            stock_item_kind = "product"
        else:
            stock_item_kind = ""
        item = {
            "tipo_item": item_type,
            "stock_item_kind": stock_item_kind,
            "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "material": str(payload.get("material", "") or "").strip(),
            "espessura": str(payload.get("espessura", "") or "").strip(),
            "operacao": str(payload.get("operacao", "") or "").strip(),
            "produto_codigo": str(payload.get("produto_codigo", "") or "").strip(),
            "produto_unid": str(payload.get("produto_unid", "") or "").strip(),
            "qtd": quantity,
            "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
            "preco_unit": round(self._parse_float(payload.get("preco_unit", 0), 0), 4),
            "desenho": str(payload.get("desenho", "") or "").strip(),
            "calc_mode": str(payload.get("calc_mode", "") or "").strip(),
            "descricao_base": str(payload.get("descricao_base", "") or "").strip(),
            "weight_total": round(self._parse_float(payload.get("weight_total", 0), 0), 3),
            "total_cost": round(self._parse_float(payload.get("total_cost", 0), 0), 2),
            "quantity_units": round(self._parse_float(payload.get("quantity_units", quantity), quantity), 2),
            "price_per_kg": round(self._parse_float(payload.get("price_per_kg", 0), 0), 4),
            "price_base_value": round(self._parse_float(payload.get("price_base_value", 0), 0), 4),
            "price_markup_pct": round(self._parse_float(payload.get("price_markup_pct", 0), 0), 2),
            "stock_metric_value": round(self._parse_float(payload.get("stock_metric_value", 0), 0), 4),
            "meters_per_unit": round(self._parse_float(payload.get("meters_per_unit", 0), 0), 3),
            "kg_per_m": round(self._parse_float(payload.get("kg_per_m", 0), 0), 4),
            "length_mm": round(self._parse_float(payload.get("length_mm", 0), 0), 1),
            "width_mm": round(self._parse_float(payload.get("width_mm", 0), 0), 1),
            "thickness_mm": round(self._parse_float(payload.get("thickness_mm", 0), 0), 2),
            "density": round(self._parse_float(payload.get("density", 0), 0), 1),
            "diameter_mm": round(self._parse_float(payload.get("diameter_mm", 0), 0), 1),
            "manual_unit_price": round(self._parse_float(payload.get("manual_unit_price", 0), 0), 4),
            "profile_section": str(payload.get("profile_section", "") or "").strip(),
            "profile_size": str(payload.get("profile_size", "") or "").strip(),
            "tube_section": str(payload.get("tube_section", "") or "").strip(),
            "quality": str(payload.get("quality", "") or "").strip(),
            "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
            "hint": str(payload.get("hint", "") or "").strip(),
            "price_base_label": str(payload.get("price_base_label", "") or "").strip(),
            "material_family": str(payload.get("material_family", "") or "").strip(),
            "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
        }
        if item_type == self.desktop_main.ORC_LINE_TYPE_PIECE:
            if not item["descricao"]:
                raise ValueError("Descricao obrigatoria na peca do conjunto.")
            if not item["material"] or not item["espessura"]:
                raise ValueError("Material e espessura sao obrigatorios nas pecas fabricadas.")
            item["produto_codigo"] = ""
            item["produto_unid"] = ""
        elif item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            product = self._product_lookup(item["produto_codigo"])
            if product is None and not item["descricao"]:
                raise ValueError("Descricao obrigatoria no produto.")
            item["_product_pending_create"] = product is None
            item["descricao"] = item["descricao"] or str((product or {}).get("descricao", "") or "").strip()
            item["produto_unid"] = item["produto_unid"] or str((product or {}).get("unid", "") or "UN").strip()
            if product is not None and item["preco_unit"] <= 0:
                item["preco_unit"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 4)
            if not item["ref_externa"]:
                item["ref_externa"] = item["produto_codigo"]
            item["material"] = ""
            item["espessura"] = ""
            item["desenho"] = ""
            item["operacao"] = item["operacao"] or "Montagem"
        else:
            if not item["descricao"]:
                raise ValueError("Descricao obrigatoria no servico de montagem.")
            item["material"] = ""
            item["espessura"] = ""
            item["produto_codigo"] = ""
            item["produto_unid"] = item["produto_unid"] or "SV"
            item["desenho"] = ""
            item["operacao"] = item["operacao"] or "Montagem"
        return item

    def assembly_model_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for model in list(self.ensure_data().get("conjuntos_modelo", []) or []):
            if not isinstance(model, dict):
                continue
            items = list(model.get("itens", []) or [])
            row = {
                "codigo": str(model.get("codigo", "") or "").strip(),
                "descricao": str(model.get("descricao", "") or "").strip(),
                "ativo": bool(model.get("ativo", True)),
                "template": bool(model.get("template", False)),
                "origem": str(model.get("origem", "") or "").strip(),
                "itens": len(items),
                "pecas": sum(1 for item in items if self.desktop_main.orc_line_is_piece(item)),
                "produtos": sum(1 for item in items if self.desktop_main.orc_line_is_product(item)),
                "servicos": sum(1 for item in items if self.desktop_main.orc_line_is_service(item)),
                "total_base": round(sum(self._parse_float(item.get("qtd", 0), 0) * self._parse_float(item.get("preco_unit", 0), 0) for item in items), 2),
                "notas": str(model.get("notas", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo", ""), item.get("descricao", "")))
        return rows

    def assembly_model_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        model = next(
            (
                row
                for row in list(self.ensure_data().get("conjuntos_modelo", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if model is None:
            raise ValueError("Conjunto nao encontrado.")
        items = [self._normalize_assembly_model_item(dict(item or {})) for item in list(model.get("itens", []) or [])]
        return {
            "codigo": str(model.get("codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "template": bool(model.get("template", False)),
            "origem": str(model.get("origem", "") or "").strip(),
            "created_at": str(model.get("created_at", "") or "").strip(),
            "updated_at": str(model.get("updated_at", "") or "").strip(),
            "itens": items,
        }

    def assembly_model_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        code = str(payload.get("codigo", "") or "").strip() or self._next_assembly_model_code()
        descricao = str(payload.get("descricao", "") or "").strip()
        if not descricao:
            raise ValueError("Descricao obrigatoria no conjunto.")
        items = [self._normalize_assembly_model_item(dict(row or {})) for row in list(payload.get("itens", []) or [])]
        if not items:
            raise ValueError("O conjunto precisa de pelo menos um item.")
        model = {
            "codigo": code,
            "descricao": descricao,
            "notas": str(payload.get("notas", "") or "").strip(),
            "ativo": bool(payload.get("ativo", True)),
            "template": bool(payload.get("template", False)),
            "origem": str(payload.get("origem", "") or "").strip(),
            "created_at": str(payload.get("created_at", "") or "").strip() or self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "itens": [{**item, "linha_ordem": index} for index, item in enumerate(items, start=1)],
        }
        existing = next(
            (
                row
                for row in list(data.get("conjuntos_modelo", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if existing is None:
            data.setdefault("conjuntos_modelo", []).append(model)
        else:
            model["created_at"] = str(existing.get("created_at", "") or "").strip() or model["created_at"]
            existing.update(model)
        self._save(force=True)
        return self.assembly_model_detail(code)

    def assembly_model_remove(self, codigo: str) -> None:
        code = str(codigo or "").strip()
        rows = list(self.ensure_data().get("conjuntos_modelo", []) or [])
        filtered = [row for row in rows if str(row.get("codigo", "") or "").strip() != code]
        if len(filtered) == len(rows):
            raise ValueError("Conjunto nao encontrado.")
        self.ensure_data()["conjuntos_modelo"] = filtered
        self._save(force=True)

    def assembly_model_expand(self, codigo: str, quantity: Any = 1) -> list[dict[str, Any]]:
        detail = self.assembly_model_detail(codigo)
        multiplier = round(self._parse_float(quantity, 0), 2)
        if multiplier <= 0:
            raise ValueError("Quantidade do conjunto invalida.")
        group_uuid = self.desktop_main.uuid.uuid4().hex[:12].upper()
        rows: list[dict[str, Any]] = []
        for item in list(detail.get("itens", []) or []):
            line = {
                "tipo_item": self.desktop_main.normalize_orc_line_type(item.get("tipo_item")),
                "stock_item_kind": str(item.get("stock_item_kind", "") or "").strip(),
                "ref_interna": "",
                "ref_externa": str(item.get("ref_externa", "") or "").strip(),
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(item.get("material", "") or "").strip(),
                "material_family": str(item.get("material_family", "") or "").strip(),
                "material_subtype": str(item.get("material_subtype", "") or "").strip(),
                "espessura": str(item.get("espessura", "") or "").strip(),
                "operacao": str(item.get("operacao", "") or "").strip(),
                "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                "conjunto_codigo": str(detail.get("codigo", "") or "").strip(),
                "conjunto_nome": str(detail.get("descricao", "") or "").strip(),
                "grupo_uuid": group_uuid,
                "qtd_base": round(self._parse_float(item.get("qtd", 0), 0), 2),
                "tempo_peca_min": round(self._parse_float(item.get("tempo_peca_min", 0), 0), 2),
                "qtd": round(self._parse_float(item.get("qtd", 0), 0) * multiplier, 2),
                "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                "desenho": str(item.get("desenho", "") or "").strip(),
                "stock_material_id": str(item.get("stock_material_id", "") or "").strip(),
                "_product_pending_create": bool(item.get("_product_pending_create", False)),
                "price_base_value": round(self._parse_float(item.get("price_base_value", 0), 0), 4),
                "price_base_label": str(item.get("price_base_label", "") or "").strip(),
                "stock_metric_value": round(self._parse_float(item.get("stock_metric_value", 0), 0), 4),
                "meters_per_unit": round(self._parse_float(item.get("meters_per_unit", 0), 0), 3),
                "kg_per_m": round(self._parse_float(item.get("kg_per_m", 0), 0), 4),
                "length_mm": round(self._parse_float(item.get("length_mm", 0), 0), 1),
                "width_mm": round(self._parse_float(item.get("width_mm", 0), 0), 1),
                "thickness_mm": round(self._parse_float(item.get("thickness_mm", 0), 0), 2),
                "diameter_mm": round(self._parse_float(item.get("diameter_mm", 0), 0), 1),
                "profile_section": str(item.get("profile_section", "") or "").strip(),
                "profile_size": str(item.get("profile_size", "") or "").strip(),
                "tube_section": str(item.get("tube_section", "") or "").strip(),
                "quality": str(item.get("quality", "") or "").strip(),
                "calc_mode": str(item.get("calc_mode", "") or "").strip(),
            }
            if self.desktop_main.orc_line_is_product(line) and not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            rows.append(line)
        return rows

    def conjunto_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for model in list(self.ensure_data().get("conjuntos", []) or []):
            if not isinstance(model, dict):
                continue
            items = list(model.get("itens", []) or [])
            row = {
                "codigo": str(model.get("codigo", "") or "").strip(),
                "descricao": str(model.get("descricao", "") or "").strip(),
                "ativo": bool(model.get("ativo", True)),
                "template": bool(model.get("template", False)),
                "origem": str(model.get("origem", "") or "").strip(),
                "itens": len(items),
                "pecas": sum(1 for item in items if self.desktop_main.orc_line_is_piece(item)),
                "produtos": sum(1 for item in items if self.desktop_main.orc_line_is_product(item)),
                "servicos": sum(1 for item in items if self.desktop_main.orc_line_is_service(item)),
                "total_custo": round(self._parse_float(model.get("total_custo", 0), 0), 2),
                "total_final": round(self._parse_float(model.get("total_final", 0), 0), 2),
                "margem_perc": round(self._parse_float(model.get("margem_perc", 0), 0), 2),
                "notas": str(model.get("notas", "") or "").strip(),
                "created_at": str(model.get("created_at", "") or "").strip(),
                "updated_at": str(model.get("updated_at", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo", ""), item.get("descricao", "")))
        return rows

    def conjunto_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        model = next(
            (
                row
                for row in list(self.ensure_data().get("conjuntos", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if model is None:
            raise ValueError("Conjunto nao encontrado.")
        items = [self._normalize_assembly_model_item(dict(item or {})) for item in list(model.get("itens", []) or [])]
        return {
            "codigo": str(model.get("codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "template": bool(model.get("template", False)),
            "origem": str(model.get("origem", "") or "").strip(),
            "margem_perc": round(self._parse_float(model.get("margem_perc", 0), 0), 2),
            "total_custo": round(self._parse_float(model.get("total_custo", 0), 0), 2),
            "total_final": round(self._parse_float(model.get("total_final", 0), 0), 2),
            "created_at": str(model.get("created_at", "") or "").strip(),
            "updated_at": str(model.get("updated_at", "") or "").strip(),
            "itens": items,
        }

    def conjunto_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        code = str(payload.get("codigo", "") or "").strip() or self._next_assembly_model_code().replace("CJM-", "CJ-")
        descricao = str(payload.get("descricao", "") or "").strip()
        if not descricao:
            raise ValueError("Descricao obrigatoria no conjunto.")
        items = [self._normalize_assembly_model_item(dict(row or {})) for row in list(payload.get("itens", []) or [])]
        if not items:
            raise ValueError("O conjunto precisa de pelo menos um item.")
        model = {
            "codigo": code,
            "descricao": descricao,
            "notas": str(payload.get("notas", "") or "").strip(),
            "ativo": bool(payload.get("ativo", True)),
            "template": bool(payload.get("template", False)),
            "origem": str(payload.get("origem", "") or "").strip(),
            "margem_perc": round(self._parse_float(payload.get("margem_perc", 0), 0), 2),
            "total_custo": round(self._parse_float(payload.get("total_custo", 0), 0), 2),
            "total_final": round(self._parse_float(payload.get("total_final", 0), 0), 2),
            "created_at": str(payload.get("created_at", "") or "").strip() or self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "itens": [{**item, "linha_ordem": index} for index, item in enumerate(items, start=1)],
        }
        existing = next(
            (
                row
                for row in list(data.get("conjuntos", []) or [])
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if existing is None:
            data.setdefault("conjuntos", []).append(model)
        else:
            model["created_at"] = str(existing.get("created_at", "") or "").strip() or model["created_at"]
            existing.update(model)
        self._save(force=True)
        return self.conjunto_detail(code)

    def conjunto_remove(self, codigo: str) -> None:
        code = str(codigo or "").strip()
        rows = list(self.ensure_data().get("conjuntos", []) or [])
        filtered = [row for row in rows if str(row.get("codigo", "") or "").strip() != code]
        if len(filtered) == len(rows):
            raise ValueError("Conjunto nao encontrado.")
        self.ensure_data()["conjuntos"] = filtered
        self._save(force=True)

    def conjunto_expand(self, codigo: str, quantity: Any = 1) -> list[dict[str, Any]]:
        detail = self.conjunto_detail(codigo)
        multiplier = round(self._parse_float(quantity, 0), 2)
        if multiplier <= 0:
            raise ValueError("Quantidade do conjunto invalida.")
        group_uuid = self.desktop_main.uuid.uuid4().hex[:12].upper()
        rows: list[dict[str, Any]] = []
        for item in list(detail.get("itens", []) or []):
            line = {
                "tipo_item": self.desktop_main.normalize_orc_line_type(item.get("tipo_item")),
                "stock_item_kind": str(item.get("stock_item_kind", "") or "").strip(),
                "ref_interna": "",
                "ref_externa": str(item.get("ref_externa", "") or "").strip(),
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(item.get("material", "") or "").strip(),
                "espessura": str(item.get("espessura", "") or "").strip(),
                "operacao": str(item.get("operacao", "") or "").strip(),
                "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                "conjunto_codigo": str(detail.get("codigo", "") or "").strip(),
                "conjunto_nome": str(detail.get("descricao", "") or "").strip(),
                "grupo_uuid": group_uuid,
                "qtd_base": round(self._parse_float(item.get("qtd", 0), 0), 2),
                "tempo_peca_min": round(self._parse_float(item.get("tempo_peca_min", 0), 0), 2),
                "qtd": round(self._parse_float(item.get("qtd", 0), 0) * multiplier, 2),
                "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                "desenho": str(item.get("desenho", "") or "").strip(),
                "stock_material_id": str(item.get("stock_material_id", "") or "").strip(),
                "_product_pending_create": bool(item.get("_product_pending_create", False)),
                "price_base_value": round(self._parse_float(item.get("price_base_value", 0), 0), 4),
                "price_base_label": str(item.get("price_base_label", "") or "").strip(),
                "stock_metric_value": round(self._parse_float(item.get("stock_metric_value", 0), 0), 4),
                "meters_per_unit": round(self._parse_float(item.get("meters_per_unit", 0), 0), 3),
                "kg_per_m": round(self._parse_float(item.get("kg_per_m", 0), 0), 4),
                "length_mm": round(self._parse_float(item.get("length_mm", 0), 0), 1),
                "width_mm": round(self._parse_float(item.get("width_mm", 0), 0), 1),
                "thickness_mm": round(self._parse_float(item.get("thickness_mm", 0), 0), 2),
                "diameter_mm": round(self._parse_float(item.get("diameter_mm", 0), 0), 1),
                "profile_section": str(item.get("profile_section", "") or "").strip(),
                "profile_size": str(item.get("profile_size", "") or "").strip(),
                "tube_section": str(item.get("tube_section", "") or "").strip(),
                "quality": str(item.get("quality", "") or "").strip(),
                "calc_mode": str(item.get("calc_mode", "") or "").strip(),
            }
            if self.desktop_main.orc_line_is_product(line) and not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            rows.append(line)
        return rows

    def _normalize_orc_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        line_type = self.desktop_main.normalize_orc_line_type(payload.get("tipo_item"))
        stock_item_kind = str(payload.get("stock_item_kind", "") or "").strip()
        if line_type == self.desktop_main.ORC_LINE_TYPE_PIECE and (
            stock_item_kind == "raw_material" or str(payload.get("stock_material_id", "") or "").strip()
        ):
            stock_item_kind = "raw_material"
        elif line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            stock_item_kind = "product"
        else:
            stock_item_kind = ""
        line = {
            "tipo_item": line_type,
            "stock_item_kind": stock_item_kind,
            "ref_interna": str(payload.get("ref_interna", "") or "").strip(),
            "ref_externa": str(payload.get("ref_externa", "") or "").strip(),
            "descricao": str(payload.get("descricao", "") or "").strip(),
            "material": str(payload.get("material", "") or "").strip(),
            "material_family": str(payload.get("material_family", "") or "").strip(),
            "material_subtype": str(payload.get("material_subtype", "") or "").strip(),
            "material_supplied_by_client": bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False)),
            "material_fornecido_cliente": bool(payload.get("material_fornecido_cliente", False) or payload.get("material_supplied_by_client", False)),
            "material_cost_included": (
                bool(payload.get("material_cost_included", True))
                if "material_cost_included" in payload
                else not bool(payload.get("material_supplied_by_client", False) or payload.get("material_fornecido_cliente", False))
            ),
            "espessura": str(payload.get("espessura", "") or "").strip(),
            "operacao": str(payload.get("operacao", "") or "").strip(),
            "produto_codigo": str(payload.get("produto_codigo", "") or "").strip(),
            "produto_unid": str(payload.get("produto_unid", "") or "").strip(),
            "conjunto_codigo": str(payload.get("conjunto_codigo", "") or "").strip(),
            "conjunto_nome": str(payload.get("conjunto_nome", "") or "").strip(),
            "grupo_uuid": str(payload.get("grupo_uuid", "") or "").strip(),
            "qtd_base": round(self._parse_float(payload.get("qtd_base", payload.get("qtd", 0)), 0), 2),
            "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0)), 0), 2),
            "qtd": round(self._parse_float(payload.get("qtd", 0), 0), 2),
            "preco_unit": round(self._parse_float(payload.get("preco_unit", 0), 0), 4),
            "desenho": str(payload.get("desenho", "") or "").strip(),
            "price_per_kg": round(self._parse_float(payload.get("price_per_kg", 0), 0), 4),
            "price_base_value": round(self._parse_float(payload.get("price_base_value", 0), 0), 4),
            "price_markup_pct": round(self._parse_float(payload.get("price_markup_pct", 0), 0), 2),
            "stock_metric_value": round(self._parse_float(payload.get("stock_metric_value", 0), 0), 4),
            "price_base_label": str(payload.get("price_base_label", "") or "").strip(),
            "kg_per_m": round(self._parse_float(payload.get("kg_per_m", 0), 0), 4),
            "meters_per_unit": round(self._parse_float(payload.get("meters_per_unit", 0), 0), 3),
            "length_mm": round(self._parse_float(payload.get("length_mm", 0), 0), 1),
            "width_mm": round(self._parse_float(payload.get("width_mm", 0), 0), 1),
            "thickness_mm": round(self._parse_float(payload.get("thickness_mm", 0), 0), 2),
            "diameter_mm": round(self._parse_float(payload.get("diameter_mm", 0), 0), 1),
            "profile_section": str(payload.get("profile_section", "") or "").strip(),
            "profile_size": str(payload.get("profile_size", "") or "").strip(),
            "tube_section": str(payload.get("tube_section", "") or "").strip(),
            "quality": str(payload.get("quality", "") or "").strip(),
            "stock_material_id": str(payload.get("stock_material_id", "") or "").strip(),
            "laser_base_active": bool(payload.get("laser_base_active", False)),
            "laser_base_tempo_unit": round(self._parse_float(payload.get("laser_base_tempo_unit", payload.get("tempo_peca_min", payload.get("tempo_pecas_min", 0))), 0), 4),
            "laser_base_preco_unit": round(self._parse_float(payload.get("laser_base_preco_unit", payload.get("preco_unit", 0)), 0), 4),
        }
        if line["material_supplied_by_client"] or line["material_fornecido_cliente"]:
            line["material_supplied_by_client"] = True
            line["material_fornecido_cliente"] = True
            line["material_cost_included"] = False
        if line["qtd"] <= 0:
            raise ValueError("Quantidade invalida na linha.")
        if line_type == self.desktop_main.ORC_LINE_TYPE_PIECE:
            if not line["descricao"]:
                raise ValueError("Descricao obrigatoria na linha.")
            if not line["material"] or not line["espessura"]:
                raise ValueError("Material e espessura sao obrigatorios na linha.")
            if not line["material_family"]:
                line["material_family"] = line["material"]
            line["produto_codigo"] = ""
            line["produto_unid"] = ""
            if stock_item_kind == "raw_material":
                line["ref_interna"] = ""
                line["desenho"] = ""
        elif line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            product = self._product_lookup(line["produto_codigo"])
            if product is None and not line["descricao"]:
                raise ValueError("Descricao obrigatoria no produto.")
            line["_product_pending_create"] = product is None
            line["descricao"] = line["descricao"] or str((product or {}).get("descricao", "") or "").strip()
            line["produto_unid"] = line["produto_unid"] or str((product or {}).get("unid", "") or "UN").strip()
            if product is not None and line["preco_unit"] <= 0:
                line["preco_unit"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 4)
            if not line["ref_externa"]:
                line["ref_externa"] = line["produto_codigo"]
            line["ref_interna"] = ""
            line["material"] = ""
            line["material_family"] = ""
            line["material_subtype"] = ""
            line["material_supplied_by_client"] = False
            line["material_fornecido_cliente"] = False
            line["material_cost_included"] = False
            line["espessura"] = ""
            line["desenho"] = ""
            line["laser_base_active"] = False
            line["laser_base_tempo_unit"] = 0.0
            line["laser_base_preco_unit"] = 0.0
            line["operacao"] = line["operacao"] or "Montagem"
        else:
            if not line["descricao"]:
                raise ValueError("Descricao obrigatoria na linha de servico.")
            line["ref_interna"] = ""
            line["material"] = ""
            line["material_family"] = ""
            line["material_subtype"] = ""
            line["material_supplied_by_client"] = False
            line["material_fornecido_cliente"] = False
            line["material_cost_included"] = False
            line["espessura"] = ""
            line["produto_codigo"] = ""
            line["produto_unid"] = line["produto_unid"] or "SV"
            line["desenho"] = ""
            line["laser_base_active"] = False
            line["laser_base_tempo_unit"] = 0.0
            line["laser_base_preco_unit"] = 0.0
            line["operacao"] = line["operacao"] or "Montagem"
        line["total"] = round(line["qtd"] * line["preco_unit"], 2)

        def _repair_laser_base(snapshot_payload: dict[str, Any]) -> None:
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE:
                return
            if not bool(line.get("laser_base_active", False)):
                return
            current_time = round(self._parse_float(line.get("tempo_peca_min", 0), 0), 4)
            current_price = round(self._parse_float(line.get("preco_unit", 0), 0), 4)
            repair_extra_time = round(
                sum(
                    self._quote_collect_non_laser_map(
                        dict(payload.get("tempos_operacao", {}) or {}),
                        dict(snapshot_payload.get("tempos_operacao", {}) or {}),
                        digits=4,
                    ).values()
                ),
                4,
            )
            repair_extra_price = round(
                sum(
                    self._quote_collect_non_laser_map(
                        dict(payload.get("custos_operacao", {}) or {}),
                        dict(snapshot_payload.get("custos_operacao", {}) or {}),
                        digits=4,
                    ).values()
                ),
                4,
            )
            max_safe_base_time = round(max(0.0, current_time - repair_extra_time), 4)
            max_safe_base_price = round(max(0.0, current_price - repair_extra_price), 4)
            if round(self._parse_float(line.get("laser_base_tempo_unit", 0), 0), 4) > max_safe_base_time + 0.0001:
                line["laser_base_tempo_unit"] = max_safe_base_time
            if round(self._parse_float(line.get("laser_base_preco_unit", 0), 0), 4) > max_safe_base_price + 0.0001:
                line["laser_base_preco_unit"] = max_safe_base_price

        def _apply_laser_base_blend(snapshot_payload: dict[str, Any]) -> None:
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE:
                return
            if not bool(line.get("laser_base_active", False)):
                return
            base_time = round(self._parse_float(line.get("laser_base_tempo_unit", 0), 0), 4)
            base_price = round(self._parse_float(line.get("laser_base_preco_unit", 0), 0), 4)
            extra_time = 0.0
            extra_price = 0.0
            for op_name, raw_value in dict(snapshot_payload.get("tempos_operacao", {}) or {}).items():
                normalized = str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if normalized and normalized != "Corte Laser":
                    extra_time += self._parse_float(raw_value, 0)
            for op_name, raw_value in dict(snapshot_payload.get("custos_operacao", {}) or {}).items():
                normalized = str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip()
                if normalized and normalized != "Corte Laser":
                    extra_price += self._parse_float(raw_value, 0)
            line["tempo_peca_min"] = round(base_time + extra_time, 2)
            line["preco_unit"] = round(base_price + extra_price, 4)
            line["total"] = round(line["qtd"] * line["preco_unit"], 2)

        snapshot_source = {**dict(payload or {}), **line}
        snapshot = self._quote_line_operation_snapshot(snapshot_source)
        _repair_laser_base(snapshot)
        _apply_laser_base_blend(snapshot)
        snapshot_source = {**dict(payload or {}), **line}
        snapshot = self._quote_line_operation_snapshot(snapshot_source)
        line["operacoes_lista"] = list(snapshot.get("operacoes", []) or [])
        line["operacoes_fluxo"] = [dict(item or {}) for item in list(snapshot.get("operacoes_fluxo", []) or []) if isinstance(item, dict)]
        line["operacoes_detalhe"] = [dict(item or {}) for item in list(snapshot.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        line["tempos_operacao"] = dict(snapshot.get("tempos_operacao", {}) or {})
        line["custos_operacao"] = dict(snapshot.get("custos_operacao", {}) or {})
        line["quote_cost_snapshot"] = dict(snapshot.get("quote_cost_snapshot", {}) or {})
        return line

    def orc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip() or self._peek_next_orc_number()
        existing = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        posto_trabalho = self._normalize_workcenter_value(payload.get("posto_trabalho", "") or (existing or {}).get("posto_trabalho", ""))
        client_payload = dict(payload.get("cliente", {}) or {})
        client = self._normalize_orc_client(client_payload)
        client_code = self._ref_client_code(client.get("codigo", ""))
        if not any(str(client.get(key, "") or "").strip() for key in ("codigo", "nome", "empresa")):
            raise ValueError("Cliente obrigatorio.")
        lines = [self._normalize_orc_line(row) for row in list(payload.get("linhas", []) or [])]
        if client_code:
            self._repair_orc_ref_history(client_code)
            taken_refs, _pairs = self._active_client_ref_usage(client_code, exclude_orc_numero=numero)
            reusable_pairs = self._known_client_ref_pairs(client_code)
            seen_refs: set[str] = set()
            seen_pairs: set[tuple[str, str]] = set()
            reserved_refs = set(taken_refs)
            for line in lines:
                if not self.desktop_main.orc_line_is_piece(line):
                    line["ref_interna"] = ""
                    continue
                if self._quote_line_is_raw_material(line):
                    line["ref_interna"] = ""
                    continue
                ref_externa = str(line.get("ref_externa", "") or "").strip()
                known_ref = self._known_client_ref_for_external(client_code, ref_externa)
                if known_ref:
                    line["ref_interna"] = known_ref
                ref_interna = str(line.get("ref_interna", "") or "").strip().upper()
                pair = (ref_externa, ref_interna)
                can_reuse_known = bool(ref_interna and pair in reusable_pairs)
                can_reuse_current = bool(ref_interna and pair in seen_pairs)
                if not ref_interna or ((ref_interna in seen_refs or ref_interna in reserved_refs) and not can_reuse_known and not can_reuse_current):
                    ref_interna = str(self.desktop_main.next_ref_interna_unique(data, client_code, list(reserved_refs | seen_refs)))
                    line["ref_interna"] = ref_interna
                seen_refs.add(ref_interna)
                seen_pairs.add((ref_externa, ref_interna))
        iva_perc = round(self._parse_float(payload.get("iva_perc", 23), 23), 2)
        desconto_perc = round(max(0.0, min(100.0, self._parse_float(payload.get("desconto_perc", (existing or {}).get("desconto_perc", 0)), 0))), 2)
        preco_transporte = round(self._parse_float(payload.get("preco_transporte", 0), 0), 2)
        custo_transporte = round(self._parse_float(payload.get("custo_transporte", (existing or {}).get("custo_transporte", 0)), 0), 2)
        paletes = round(self._parse_float(payload.get("paletes", (existing or {}).get("paletes", 0)), 0), 2)
        peso_bruto_kg = round(self._parse_float(payload.get("peso_bruto_kg", (existing or {}).get("peso_bruto_kg", 0)), 0), 2)
        volume_m3 = round(self._parse_float(payload.get("volume_m3", (existing or {}).get("volume_m3", 0)), 0), 3)
        transportadora_id, transportadora_nome, _transportadora_contacto = self._normalize_supplier_reference(
            payload.get("transportadora_id", (existing or {}).get("transportadora_id", "")),
            payload.get("transportadora_nome", (existing or {}).get("transportadora_nome", "")),
        )
        referencia_transporte = str(payload.get("referencia_transporte", (existing or {}).get("referencia_transporte", "")) or "").strip()
        zona_transporte = str(payload.get("zona_transporte", (existing or {}).get("zona_transporte", "")) or "").strip()
        subtotal_linhas = round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2)
        subtotal_bruto = round(subtotal_linhas + preco_transporte, 2)
        desconto_valor = round(subtotal_bruto * (desconto_perc / 100.0), 2)
        subtotal = round(max(0.0, subtotal_bruto - desconto_valor), 2)
        total = round(subtotal * (1.0 + (iva_perc / 100.0)), 2)
        note = {
            "numero": numero,
            "data": str(payload.get("data", "") or existing.get("data", "") if isinstance(existing, dict) else "") or self.desktop_main.now_iso(),
            "estado": str(payload.get("estado", "") or (existing or {}).get("estado", "") or "Em edição"),
            "cliente": client,
            "posto_trabalho": posto_trabalho,
            "linhas": lines,
            "iva_perc": iva_perc,
            "desconto_perc": desconto_perc,
            "desconto_valor": desconto_valor,
            "preco_transporte": preco_transporte,
            "custo_transporte": custo_transporte,
            "paletes": paletes,
            "peso_bruto_kg": peso_bruto_kg,
            "volume_m3": volume_m3,
            "transportadora_id": transportadora_id,
            "transportadora_nome": transportadora_nome,
            "referencia_transporte": referencia_transporte,
            "zona_transporte": zona_transporte,
            "subtotal_linhas": subtotal_linhas,
            "subtotal_bruto": subtotal_bruto,
            "subtotal": subtotal,
            "total": total,
            "numero_encomenda": str(payload.get("numero_encomenda", "") or (existing or {}).get("numero_encomenda", "") or "").strip(),
            "ano": int(str(payload.get("ano", "") or (existing or {}).get("ano", "") or self.desktop_main.datetime.now().year)),
            "executado_por": str(payload.get("executado_por", "") or (existing or {}).get("executado_por", "") or "").strip(),
            "nota_transporte": str(payload.get("nota_transporte", "") or (existing or {}).get("nota_transporte", "") or "").strip(),
            "notas_pdf": str(payload.get("notas_pdf", "") or (existing or {}).get("notas_pdf", "") or "").strip(),
            "nota_cliente": str(payload.get("nota_cliente", "") or (existing or {}).get("nota_cliente", "") or "").strip(),
        }
        if existing is None:
            data.setdefault("orcamentos", []).append(note)
            if numero == self._peek_next_orc_number():
                try:
                    data["orc_seq"] = max(int(data.get("orc_seq", 1) or 1), int(numero.rsplit("-", 1)[-1]) + 1)
                except Exception:
                    pass
        else:
            existing.update(note)
            note = existing
        self._sync_quote_piece_registry(note)
        self._save(force=True)
        return self.orc_detail(numero)

    def orc_remove(self, numero: str) -> None:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        before = len(list(data.get("orcamentos", []) or []))
        data["orcamentos"] = [row for row in list(data.get("orcamentos", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        if len(data["orcamentos"]) == before:
            raise ValueError("Or?amento n?o encontrado.")
        self._save(force=True)
        try:
            self._mysql_delete_orc_nesting_studies(numero)
        except Exception:
            pass

    def orc_set_state(self, numero: str, estado: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        orc["estado"] = str(estado or "").strip() or "Em edição"
        self._sync_quote_piece_registry(orc)
        self._save(force=True)
        return self.orc_detail(numero)

    def _orc_render_helper(self) -> Any:
        helper = SimpleNamespace(data=self.ensure_data())
        helper._extract_orc_operacoes = lambda orc=None: self.orc_actions._extract_orc_operacoes(helper, orc)
        helper._build_orc_notes_lines = lambda orc: self.orc_actions._build_orc_notes_lines(helper, orc)
        return helper

    def orc_render_pdf(self, numero: str, path: str | Path) -> Path:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        target = Path(path)
        helper = self._orc_render_helper()
        self.orc_actions.render_orc_pdf(helper, str(target), orc)
        return target

    def orc_open_pdf(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_orcamento_{str(numero or '').strip()}.pdf"
        self.orc_render_pdf(numero, target)
        os.startfile(str(target))
        return target

    def orc_render_nesting_study_pdf(self, numero: str, path: str | Path, group_key: str = "") -> Path:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        numero_txt = str(numero or "").strip()
        detail = self.orc_detail(numero_txt)
        studies = self.orc_nesting_studies(numero_txt)
        if not studies:
            raise ValueError("Este orçamento ainda não tem estudos de nesting guardados.")
        selected_key = str(group_key or detail.get("nesting_group_key", "") or "").strip()
        if not selected_key or selected_key not in studies:
            ordered_keys = sorted(studies.keys())
            selected_key = ordered_keys[0]
        study = dict(studies.get(selected_key, {}) or {})
        if not study:
            raise ValueError("Estudo de nesting não encontrado.")

        result_data = dict(study.get("result_data", {}) or {})
        summary = dict(result_data.get("summary", study.get("summary", {})) or {})
        bridge = dict(study.get("quote_bridge", {}) or {})
        cost_report = dict(study.get("cost_report", {}) or {})
        options = dict(study.get("options", {}) or {})
        sheets = [dict(row or {}) for row in list(result_data.get("sheets", []) or [])]
        sheet_candidates = [dict(row or {}) for row in list(result_data.get("sheet_candidates", []) or [])]
        unplaced = [dict(row or {}) for row in list(result_data.get("unplaced", []) or [])]
        warnings = [str(row or "").strip() for row in list(result_data.get("warnings", []) or []) if str(row or "").strip()]
        part_rows = [dict(row or {}) for row in list(cost_report.get("part_rows", []) or bridge.get("part_rows", [])) if isinstance(row, dict)]
        decision_lines = [str(row or "").strip() for row in list(cost_report.get("decision_lines", []) or []) if str(row or "").strip()]
        totals = dict(cost_report.get("totals", {}) or {})

        target = Path(path)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        page_width, page_height = landscape(A4)
        margin = 26
        c = pdf_canvas.Canvas(str(target), pagesize=landscape(A4))

        font_regular = "Helvetica"
        font_bold = "Helvetica-Bold"
        for name, file_name in (("SegoeUI", "segoeui.ttf"), ("SegoeUI-Bold", "segoeuib.ttf")):
            font_path = Path(r"C:\Windows\Fonts") / file_name
            if font_path.exists():
                try:
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont

                    pdfmetrics.registerFont(TTFont(name, str(font_path)))
                    if "Bold" in name:
                        font_bold = name
                    else:
                        font_regular = name
                except Exception:
                    pass

        def set_font(bold: bool, size: float) -> None:
            c.setFont(font_bold if bold else font_regular, size)

        def draw_header(title: str, subtitle: str) -> float:
            c.setFillColor(palette["primary"])
            c.roundRect(margin, page_height - margin - 60, page_width - (margin * 2), 60, 18, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 18)
            c.drawString(margin + 16, page_height - margin - 24, title)
            set_font(False, 9)
            c.drawString(margin + 16, page_height - margin - 40, subtitle)
            company = str(branding.get("company_name", "") or "luGEST").strip() or "luGEST"
            generated = datetime.now().strftime("%d/%m/%Y %H:%M")
            c.drawRightString(page_width - margin - 16, page_height - margin - 24, company)
            c.drawRightString(page_width - margin - 16, page_height - margin - 40, generated)
            return page_height - margin - 74

        def draw_footer() -> None:
            c.setFillColor(palette["muted"])
            set_font(False, 8)
            c.drawString(margin, 20, "Estudo de nesting guardado por orçamento, ligado ao Plano de Chapa e ao custo do lote.")
            c.drawRightString(page_width - margin, 20, f"Orçamento {numero_txt}")

        def draw_metric_card(x: float, y_top: float, width: float, title: str, value: str, accent: Any) -> None:
            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(x, y_top - 52, width, 46, 12, stroke=1, fill=1)
            c.setFillColor(accent)
            c.rect(x + 10, y_top - 20, 26, 4, stroke=0, fill=1)
            c.setFillColor(palette["muted"])
            set_font(True, 8.2)
            c.drawString(x + 10, y_top - 14, title)
            c.setFillColor(palette["ink"])
            set_font(True, 13)
            c.drawString(x + 10, y_top - 35, value)

        def draw_info_box(x: float, y_top: float, width: float, title: str, lines: list[str], *, tone: str = "default") -> float:
            body_lines = [line for line in lines if str(line or "").strip()]
            wrapped: list[str] = []
            for line in body_lines:
                wrapped.extend(_pdf_wrap_text(line, font_regular, 8.4, width - 22, max_lines=3) or ["-"])
            box_height = max(54.0, 18.0 + (len(wrapped) * 11.0) + 18.0)
            tone_fill = palette["primary_soft_2"] if tone == "info" else colors.HexColor("#FFF8EB") if tone == "warning" else colors.white
            c.setFillColor(tone_fill)
            c.setStrokeColor(palette["line"])
            c.roundRect(x, y_top - box_height, width, box_height, 12, stroke=1, fill=1)
            c.setFillColor(palette["ink"])
            set_font(True, 10)
            c.drawString(x + 10, y_top - 16, title)
            set_font(False, 8.4)
            cursor_y = y_top - 30
            for line in wrapped:
                c.drawString(x + 10, cursor_y, line)
                cursor_y -= 11
            return box_height

        def ensure_page(current_y: float, needed: float, title: str, subtitle: str) -> float:
            if current_y - needed >= 46:
                return current_y
            draw_footer()
            c.showPage()
            return draw_header(title, subtitle)

        def draw_table_header(y_top: float, columns: list[tuple[str, float]]) -> tuple[float, list[float], float]:
            total_width = page_width - (margin * 2)
            c.setFillColor(palette["primary_dark"])
            c.roundRect(margin, y_top - 20, total_width, 18, 8, stroke=0, fill=1)
            c.setFillColor(colors.white)
            set_font(True, 8)
            x_positions: list[float] = []
            cursor_x = margin + 7
            for label, ratio in columns:
                x_positions.append(cursor_x)
                c.drawString(cursor_x, y_top - 13, label)
                cursor_x += total_width * ratio
            return y_top - 24, x_positions, total_width

        def draw_sheet_map(x: float, y_top: float, width: float, height: float, sheet: dict[str, Any]) -> None:
            def draw_polygon(points: list[tuple[float, float]], *, stroke_color: Any, fill_color: Any | None = None, stroke_width: float = 1.0) -> None:
                if len(points) < 3:
                    return
                path = c.beginPath()
                path.moveTo(points[0][0], points[0][1])
                for px, py in points[1:]:
                    path.lineTo(px, py)
                path.close()
                c.setStrokeColor(stroke_color)
                c.setLineWidth(stroke_width)
                if fill_color is not None:
                    c.setFillColor(fill_color)
                    c.drawPath(path, stroke=1, fill=1)
                else:
                    c.drawPath(path, stroke=1, fill=0)

            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(x, y_top - height, width, height, 14, stroke=1, fill=1)
            c.setFillColor(palette["primary_soft_2"])
            c.roundRect(x + 8, y_top - 26, width - 16, 10, 6, stroke=0, fill=1)
            title = (
                f"Chapa {int(sheet.get('index', 0) or 0)} | "
                f"{str(sheet.get('source_label', summary.get('selected_sheet_profile', {}).get('name', '-')) or '-').strip()}"
            )
            c.setFillColor(palette["ink"])
            set_font(True, 11.2)
            c.drawString(x + 10, y_top - 16, _pdf_clip_text(title, width - 20, font_bold, 9.2))
            set_font(False, 8.2)
            c.setFillColor(palette["muted"])
            tech_line = (
                f"{self._fmt(float(sheet.get('sheet_width_mm', 0) or 0))} x "
                f"{self._fmt(float(sheet.get('sheet_height_mm', 0) or 0))} mm | "
                f"origem {str(sheet.get('source_kind', '-') or '-').strip()} | "
                f"real {self._fmt(sheet.get('utilization_net_pct', 0))}%"
            )
            c.drawString(
                x + 10,
                y_top - 28,
                _pdf_clip_text(tech_line, width - 20, font_regular, 7.8),
            )
            c.drawString(
                x + 10,
                y_top - 40,
                f"{int(sheet.get('part_count', 0) or 0)} peca(s) | bbox {self._fmt(sheet.get('utilization_bbox_pct', 0))}% | compra {'sim' if str(sheet.get('source_kind', '') or '').strip().lower() == 'purchase' else 'nao'}",
            )
            legend_rows = [dict(row or {}) for row in list(sheet.get("placements", []) or [])]
            inner_x = x + 12
            inner_y_top = y_top - 52
            inner_h = height - 66
            legend_w = max(154.0, min(210.0, width * 0.25))
            panel_gap = 10.0
            map_x = inner_x + legend_w + panel_gap
            map_w = width - 24 - legend_w - panel_gap
            map_y = y_top - height + 14
            map_h = max(80.0, inner_h)
            c.setFillColor(colors.HexColor("#F9FBFE"))
            c.roundRect(inner_x, map_y, legend_w, map_h, 12, stroke=0, fill=1)
            c.setStrokeColor(palette["line"])
            c.roundRect(inner_x, map_y, legend_w, map_h, 12, stroke=1, fill=0)
            c.setFillColor(colors.HexColor("#F8FAFC"))
            c.roundRect(map_x, map_y, map_w, map_h, 12, stroke=0, fill=1)
            c.setStrokeColor(colors.HexColor("#9cb2c8"))
            c.roundRect(map_x, map_y, map_w, map_h, 12, stroke=1, fill=0)

            set_font(True, 8.4)
            c.setFillColor(palette["ink"])
            c.drawString(inner_x + 10, y_top - 66, "Peças no layout")
            set_font(False, 7.2)
            legend_y = y_top - 80
            max_legend_rows = min(len(legend_rows), 16)
            for idx, placement in enumerate(legend_rows[:max_legend_rows], start=1):
                ref_txt = str(placement.get("ref_externa", "") or placement.get("description", "") or "-").strip() or "-"
                prefix = f"{idx:02d}"
                c.setFillColor(colors.HexColor("#0f172a"))
                c.drawString(inner_x + 10, legend_y, prefix)
                c.drawString(
                    inner_x + 28,
                    legend_y,
                    _pdf_clip_text(ref_txt, legend_w - 40, font_regular, 7.2),
                )
                legend_y -= 10
            remaining = len(legend_rows) - max_legend_rows
            if remaining > 0:
                c.setFillColor(palette["muted"])
                c.drawString(inner_x + 10, legend_y, f"+{remaining} peça(s) adicionais")

            sheet_w = max(1.0, float(sheet.get("sheet_width_mm", 0) or 1.0))
            sheet_h = max(1.0, float(sheet.get("sheet_height_mm", 0) or 1.0))
            draw_pad = 10.0
            scale = min((map_w - draw_pad * 2.0) / sheet_w, (map_h - draw_pad * 2.0) / sheet_h)
            body_w = sheet_w * scale
            body_h = sheet_h * scale
            offset_x = map_x + ((map_w - body_w) / 2.0)
            offset_y = map_y + ((map_h - body_h) / 2.0)

            def map_point(px: float, py: float) -> tuple[float, float]:
                return (offset_x + (px * scale), offset_y + (py * scale))

            outer_polygons = [list(points or []) for points in list(sheet.get("sheet_outer_polygons", []) or [])]
            hole_polygons = [list(points or []) for points in list(sheet.get("sheet_hole_polygons", []) or [])]
            c.setStrokeColor(palette["line_strong"])
            c.setFillColor(colors.HexColor("#F8FAFC"))
            if outer_polygons:
                for polygon_points in outer_polygons:
                    mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                    if len(mapped) >= 3:
                        draw_polygon(mapped, stroke_color=palette["line_strong"], fill_color=colors.white, stroke_width=1.1)
                for polygon_points in hole_polygons:
                    mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                    if len(mapped) >= 3:
                        draw_polygon(mapped, stroke_color=palette["line_strong"], fill_color=colors.white, stroke_width=0.9)
            else:
                c.rect(offset_x, offset_y, body_w, body_h, stroke=1, fill=1)

            palette_hexes = ["#dbeafe", "#dcfce7", "#fef3c7", "#ffe4e6", "#ede9fe", "#cffafe", "#e2e8f0"]
            for idx, placement in enumerate(list(sheet.get("placements", []) or []), start=1):
                c.setStrokeColor(colors.HexColor("#274c77"))
                c.setFillColor(colors.HexColor(palette_hexes[idx % len(palette_hexes)]))
                poly_groups = [list(points or []) for points in list(placement.get("shape_outer_polygons", []) or [])]
                label_x = offset_x + (float(placement.get("x_mm", 0) or 0) * scale)
                label_y = offset_y + (float(placement.get("y_mm", 0) or 0) * scale)
                label_w = max(1.0, float(placement.get("width_mm", 0) or 0) * scale)
                label_h = max(1.0, float(placement.get("height_mm", 0) or 0) * scale)
                if poly_groups:
                    for polygon_points in poly_groups:
                        mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                        if len(mapped) >= 3:
                            draw_polygon(mapped, stroke_color=colors.HexColor("#274c77"), fill_color=colors.HexColor(palette_hexes[idx % len(palette_hexes)]), stroke_width=0.9)
                else:
                    c.rect(
                        label_x,
                        label_y,
                        label_w,
                        label_h,
                        stroke=1,
                        fill=1,
                    )
                if label_w >= 18 and label_h >= 12:
                    c.setFillColor(colors.HexColor("#0f172a"))
                    set_font(True, 6.8)
                    c.drawCentredString(label_x + (label_w / 2.0), label_y + (label_h / 2.0), str(idx))

        group_label = str(study.get("group_label", "") or selected_key).strip() or selected_key
        profile_name = str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or bridge.get("selected_profile_name", "") or "Apenas stock").strip() or "Apenas stock"
        subtitle = f"Orçamento {numero_txt} | Grupo {group_label} | Perfil {profile_name}"
        y = draw_header("Estudo de Nesting + Custo", subtitle)

        cards = [
            ("Programadas", f"{int(bridge.get('part_count_placed', summary.get('part_count_placed', 0)) or 0)}/{int(bridge.get('part_count_requested', summary.get('part_count_requested', 0)) or 0)}", palette["primary"]),
            ("Chapas", str(int(bridge.get("sheet_count", summary.get("sheet_count", 0)) or 0)), palette["success"]),
            ("Util. real", f"{self._fmt(summary.get('utilization_net_pct', 0))}%", palette["warning"]),
            ("Matéria", self._fmt_eur(summary.get("material_net_cost_eur", 0)), palette["primary_dark"]),
            ("Compra", self._fmt_eur(summary.get("material_purchase_requirement_eur", 0)), palette["danger"]),
        ]
        card_gap = 8
        card_width = (page_width - (margin * 2) - (card_gap * (len(cards) - 1))) / len(cards)
        card_x = margin
        for title, value, accent in cards:
            draw_metric_card(card_x, y, card_width, title, value, accent)
            card_x += card_width + card_gap
        y -= 64

        study_lines = [
            f"Cliente: {str(dict(detail.get('cliente', {}) or {}).get('nome', '') or '-').strip() or '-'}",
            f"Método: {str(bridge.get('analysis_method', summary.get('selection_mode', '-')) or '-').strip() or '-'}",
            f"Regras: margem peça {float(options.get('part_spacing_mm', 0) or 0):.1f} mm | margem borda {float(options.get('edge_margin_mm', 0) or 0):.1f} mm | rotação auto {'sim' if bool(options.get('allow_rotate')) else 'não'}",
            f"Fluxo: stock primeiro {'sim' if bool(options.get('use_stock_first')) else 'não'} | compra complementar {'sim' if bool(options.get('allow_purchase_fallback', True)) else 'não'} | contorno {'sim' if bool(options.get('shape_aware')) else 'não'}",
        ]
        report_lines = [
            f"Valor comercial colocado: {self._fmt_eur(totals.get('quoted_total_eur', bridge.get('quoted_total_eur', 0)))}",
            f"Tempo máquina: {self._fmt(totals.get('machine_total_min', 0))} min | corte {self._fmt(totals.get('cut_length_m', 0))} m | pierces {int(totals.get('pierce_count', 0) or 0)}",
            f"Stock usado: {int(summary.get('stock_sheet_count', 0) or 0)} | retalhos {int(summary.get('remnant_sheet_count', 0) or 0)} | compra {int(summary.get('purchased_sheet_count', 0) or 0)}",
        ]
        left_box_h = draw_info_box(margin, y, (page_width - (margin * 2) - 10) * 0.52, "Contexto do estudo", study_lines, tone="info")
        right_box_h = draw_info_box(margin + ((page_width - (margin * 2) - 10) * 0.52) + 10, y, (page_width - (margin * 2) - 10) * 0.48, "Resumo económico", report_lines, tone="default")
        y -= max(left_box_h, right_box_h) + 10

        note_lines = (decision_lines[:4] or warnings[:4] or ["Sem observações adicionais registadas."])
        note_height = draw_info_box(margin, y, page_width - (margin * 2), "Decisão e observações", note_lines, tone="warning" if warnings else "default")
        y -= note_height + 12

        candidate_columns = [
            ("Cenário", 0.29),
            ("Método", 0.29),
            ("Chapas", 0.09),
            ("Compact.", 0.11),
            ("Compra m2", 0.11),
            ("Total m2", 0.11),
        ]
        if sheet_candidates:
            y = ensure_page(y, 90, "Estudo de Nesting + Custo", subtitle)
            y, x_positions, total_w = draw_table_header(y, candidate_columns)
            row_h = 18
            for row_index, candidate in enumerate(sheet_candidates[:8]):
                y = ensure_page(y, row_h + 8, "Estudo de Nesting + Custo", subtitle)
                fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
                c.setFillColor(fill_color)
                c.setStrokeColor(palette["line"])
                c.roundRect(margin, y - row_h + 2, page_width - (margin * 2), row_h - 2, 8, stroke=1, fill=1)
                values = [
                    str(candidate.get("name", "") or "-").strip() or "-",
                    str(candidate.get("method", "") or bridge.get("analysis_method", "-")).strip() or "-",
                    str(int(candidate.get("sheet_count", 0) or 0)),
                    f"{self._fmt(candidate.get('layout_compactness_pct', 0))}%",
                    f"{self._fmt((candidate.get('purchase_sheet_area_mm2', 0) or 0) / 1_000_000.0)}",
                    f"{self._fmt((candidate.get('total_sheet_area_mm2', 0) or 0) / 1_000_000.0)}",
                ]
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                for idx, value in enumerate(values):
                    max_w = (total_w * candidate_columns[idx][1]) - 10
                    draw_value = _pdf_clip_text(value, max_w, font_regular, 8)
                    c.drawString(x_positions[idx], y - 10, draw_value)
                y -= row_h
            y -= 8

        part_columns = [
            ("Ref.", 0.16),
            ("Descrição", 0.34),
            ("Qtd", 0.07),
            ("Tempo", 0.10),
            ("Corte", 0.10),
            ("Pierces", 0.09),
            ("Valor", 0.14),
        ]
        if part_rows:
            y = ensure_page(y, 110, "Estudo de Nesting + Custo", subtitle)
            y, x_positions, total_w = draw_table_header(y, part_columns)
            row_h = 18
            for row_index, row in enumerate(part_rows):
                y = ensure_page(y, row_h + 8, "Estudo de Nesting + Custo", subtitle)
                fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
                c.setFillColor(fill_color)
                c.setStrokeColor(palette["line"])
                c.roundRect(margin, y - row_h + 2, page_width - (margin * 2), row_h - 2, 8, stroke=1, fill=1)
                values = [
                    str(row.get("ref_externa", "") or "-").strip() or "-",
                    str(row.get("description", "") or "-").strip() or "-",
                    str(int(row.get("qty", 0) or 0)),
                    f"{self._fmt(row.get('machine_total_min', 0))} min",
                    f"{self._fmt(row.get('cut_length_m', 0))} m",
                    str(int(row.get("pierce_count", 0) or 0)),
                    self._fmt_eur(row.get("quoted_total_eur", 0)),
                ]
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                for idx, value in enumerate(values):
                    max_w = (total_w * part_columns[idx][1]) - 10
                    draw_value = _pdf_clip_text(value, max_w, font_regular, 8)
                    c.drawString(x_positions[idx], y - 10, draw_value)
                y -= row_h
            y -= 10

        if unplaced:
            y = ensure_page(y, 80, "Estudo de Nesting + Custo", subtitle)
            unplaced_preview = [
                f"{str(row.get('ref_externa', '-') or '-').strip() or '-'} | {str(row.get('description', '-') or '-').strip() or '-'}"
                for row in unplaced[:6]
            ]
            box_h = draw_info_box(margin, y, page_width - (margin * 2), "Peças fora do plano", unplaced_preview, tone="warning")
            y -= box_h + 10

        if sheets:
            draw_footer()
            c.showPage()
            y = draw_header("Mapas de Chapa", subtitle)
            box_width = page_width - (margin * 2)
            box_height = page_height - margin - 74 - 36
            current_y = y
            for sheet in sheets:
                if current_y - box_height < 50:
                    draw_footer()
                    c.showPage()
                    current_y = draw_header("Mapas de Chapa", subtitle)
                draw_sheet_map(margin, current_y, box_width, box_height, sheet)
                current_y -= box_height + 14
                if current_y - box_height >= 50:
                    draw_footer()
                    c.showPage()
                    current_y = draw_header("Mapas de Chapa", subtitle)

        draw_footer()
        c.save()
        return target

    def orc_open_nesting_study_pdf(self, numero: str, group_key: str = "") -> Path:
        safe_group = "".join(ch if ch.isalnum() else "_" for ch in str(group_key or "").strip()) or "grupo"
        target = Path(tempfile.gettempdir()) / f"lugest_nesting_{str(numero or '').strip()}_{safe_group}.pdf"
        self.orc_render_nesting_study_pdf(numero, target, group_key=group_key)
        os.startfile(str(target))
        return target

    def _quote_line_is_production_ready(self, line: dict[str, Any] | None = None) -> bool:
        row = dict(line or {})
        if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        if self._quote_line_is_raw_material(row):
            return False
        drawing_path = str(row.get("desenho", "") or "").strip()
        ops = [
            str(self.desktop_main.normalize_operacao_nome(op) or op or "").strip()
            for op in self._planning_ops_from_ops_value(row.get("operacao", ""))
        ]
        ops = [op for op in ops if op and op != "Montagem"]
        material = str(row.get("material", "") or "").strip()
        thickness = str(row.get("espessura", "") or "").strip()
        time_per_piece = self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", 0)), 0)
        detail_ready = bool(
            list(row.get("operacoes_detalhe", []) or [])
            or dict(row.get("tempos_operacao", {}) or {})
            or dict(row.get("custos_operacao", {}) or {})
        )
        has_work = bool(ops or detail_ready or time_per_piece > 0)
        return bool(has_work and (drawing_path or (material and thickness)))

    def _quote_line_is_raw_material(self, line: dict[str, Any] | None = None) -> bool:
        row = dict(line or {})
        if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        if str(row.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(row.get("stock_material_id", "") or "").strip():
            return True
        subtype = self.desktop_main.norm_text(str(row.get("material_subtype", "") or row.get("calc_mode", "") or "").strip())
        if subtype == "stockmp":
            return True
        return False

    def _quote_line_production_route(self, line: dict[str, Any] | None = None) -> str:
        row = dict(line or {})
        line_type = self.desktop_main.normalize_orc_line_type(row.get("tipo_item"))
        if line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
            return "montagem"
        if line_type == self.desktop_main.ORC_LINE_TYPE_SERVICE:
            return "montagem"
        if not self._quote_line_is_production_ready(row):
            return "conjunto"
        ops = [str(self.desktop_main.normalize_operacao_nome(op) or op or "").strip() for op in self._planning_ops_from_ops_value(row.get("operacao", ""))]
        ops = [op for op in ops if op]
        subtype_norm = self.desktop_main.norm_text(str(row.get("material_subtype", "") or row.get("calc_mode", "") or "").strip())
        material_norm = self.desktop_main.norm_text(str(row.get("material", "") or "").strip())
        drawing_path = str(row.get("desenho", "") or "").strip()
        if "Corte Laser" in ops and drawing_path:
            return "laser"
        if any(token in subtype_norm for token in ("tubo", "cantoneira", "perfil", "barra", "ferronervurado", "chapa")):
            return "serralharia"
        if any(token in material_norm for token in ("tubo", "cantoneira", "perfil", "barra", "ferro nervurado")):
            return "serralharia"
        if "Serralharia" in ops:
            return "serralharia"
        if "Corte Laser" in ops:
            return "laser"
        if "Montagem" in ops:
            return "montagem"
        return "conjunto"

    def orc_convert_to_order(self, numero: str, nota_cliente: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        note = str(nota_cliente or "").strip()
        orc = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Or?amento n?o encontrado.")
        if str(orc.get("numero_encomenda", "") or "").strip():
            raise ValueError("Orcamento ja convertido.")
        estado_norm = str(orc.get("estado", "") or "").strip().lower()
        if "aprovado" not in estado_norm:
            raise ValueError("Apenas orcamentos aprovados podem ser convertidos.")
        if not list(orc.get("linhas", []) or []):
            raise ValueError("Sem linhas para converter.")
        cli = self._normalize_orc_client(orc.get("cliente", {}))
        codigo = str(cli.get("codigo", "") or "").strip()
        if codigo and self.desktop_main.find_cliente(data, codigo):
            cliente_code = codigo
        else:
            cliente_code = ""
            for row in list(data.get("clientes", []) or []):
                if not isinstance(row, dict):
                    continue
                if cli.get("nif") and str(row.get("nif", "") or "").strip() == str(cli.get("nif", "") or "").strip():
                    cliente_code = str(row.get("codigo", "") or "").strip()
                    break
                if cli.get("nome") and str(row.get("nome", "") or "").strip() == str(cli.get("nome", "") or "").strip():
                    cliente_code = str(row.get("codigo", "") or "").strip()
                    break
            if not cliente_code:
                cliente_code = str(self.desktop_main.next_cliente_codigo(data))
                data.setdefault("clientes", []).append(
                    {
                        "codigo": cliente_code,
                        "nome": str(cli.get("nome", "") or "").strip(),
                        "nif": str(cli.get("nif", "") or "").strip(),
                        "morada": str(cli.get("morada", "") or "").strip(),
                        "contacto": str(cli.get("contacto", "") or "").strip(),
                        "email": str(cli.get("email", "") or "").strip(),
                        "observacoes": "",
                    }
                )
        alert_txt = (
            f"ALERTA: Encomenda gerada por conversao do orcamento {orc.get('numero')}. "
            "Confirmar dados de cliente, materiais, espessuras e prazos."
        )
        obs_txt = f"{alert_txt} | Origem: Orcamento {orc.get('numero')}"
        if note:
            obs_txt += f" | Nota cliente: {note}"
        enc = {
            "numero": self.desktop_main.next_encomenda_numero(data),
            "cliente": cliente_code,
            "nota_cliente": note,
            "nota_transporte": str(orc.get("nota_transporte", "") or "").strip(),
            "preco_transporte": round(self._parse_float(orc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(orc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self._parse_float(orc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(orc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(orc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(orc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(orc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(orc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": str(orc.get("zona_transporte", "") or "").strip(),
            "local_descarga": str(cli.get("morada", "") or "").strip(),
            "transporte_numero": "",
            "estado_transporte": "",
            "data_criacao": self.desktop_main.now_iso(),
            "data_entrega": "",
            "tempo": 0.0,
            "tempo_estimado": 0.0,
            "cativar": False,
            "posto_trabalho": self._normalize_workcenter_value(orc.get("posto_trabalho", "")),
            "observacoes": obs_txt,
            "alerta_conversao": True,
            "estado": "Preparacao",
            "materiais": [],
            "reservas": [],
            "montagem_itens": [],
            "numero_orcamento": orc.get("numero"),
        }
        mats: dict[str, dict[str, Any]] = {}
        piece_idx = 1
        total_time = 0.0
        used_refs: set[str] = set()
        montagem_items: list[dict[str, Any]] = []
        for line in list(orc.get("linhas", []) or []):
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            production_route = self._quote_line_production_route(line)
            qtd_line = float(line.get("qtd", 0) or 0)
            tempo_peca = float(line.get("tempo_peca_min", line.get("tempo_pecas_min", 0)) or 0)
            total_time += tempo_peca * max(qtd_line, 0.0)
            if production_route in {"montagem", "conjunto"}:
                montagem_items.append(
                    {
                        "linha_ordem": len(montagem_items) + 1,
                        "tipo_item": line_type,
                        "stock_item_kind": str(line.get("stock_item_kind", "") or "").strip(),
                        "descricao": str(line.get("descricao", "") or "").strip(),
                        "material": str(line.get("material", "") or "").strip(),
                        "material_family": str(line.get("material_family", "") or "").strip(),
                        "material_subtype": str(line.get("material_subtype", "") or "").strip(),
                        "espessura": str(line.get("espessura", "") or "").strip(),
                        "stock_material_id": str(line.get("stock_material_id", "") or "").strip(),
                        "produto_codigo": str(line.get("produto_codigo", "") or "").strip(),
                        "produto_unid": str(line.get("produto_unid", "") or "").strip(),
                        "_product_pending_create": bool(line.get("_product_pending_create", False)),
                        "qtd_planeada": round(qtd_line, 2),
                        "qtd_consumida": 0.0,
                        "preco_unit": round(self._parse_float(line.get("preco_unit", 0), 0), 4),
                        "conjunto_codigo": str(line.get("conjunto_codigo", "") or "").strip(),
                        "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                        "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                        "estado": "Pendente" if production_route == "montagem" else "Componente",
                        "obs": (
                            str(line.get("operacao", "") or production_route).strip()
                            if production_route == "montagem"
                            else "Conjunto sem desenho/operacoes tecnicas. Nao segue para operador."
                        ),
                        "created_at": self.desktop_main.now_iso(),
                        "consumed_at": "",
                        "consumed_by": "",
                    }
                )
                continue
            material = str(line.get("material", "") or "").strip()
            espessura = str(line.get("espessura", "") or "").strip()
            if not material or not espessura:
                raise ValueError("Todas as linhas precisam de material e espessura.")
            mats.setdefault(material, {"material": material, "estado": "Preparacao", "espessuras": {}})
            mats[material]["espessuras"].setdefault(
                espessura,
                {"espessura": espessura, "tempo_min": 0.0, "tempos_operacao": {}, "maquinas_operacao": {}, "estado": "Preparacao", "pecas": []},
            )
            planning_ops = [op for op in self._planning_ops_from_ops_value(line.get("operacao", "")) if op != "Montagem"]
            if production_route == "serralharia":
                planning_ops = [op for op in planning_ops if op != "Corte Laser"]
                if "Serralharia" not in planning_ops:
                    planning_ops.insert(0, "Serralharia")
            elif production_route == "laser":
                if "Corte Laser" not in planning_ops:
                    planning_ops.insert(0, "Corte Laser")
            esp_bucket = mats[material]["espessuras"][espessura]
            tempos_operacao = esp_bucket.setdefault("tempos_operacao", {})
            maquinas_operacao = esp_bucket.setdefault("maquinas_operacao", {})
            detailed_op_times = {
                str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip(): self._parse_float(raw_value, 0)
                for op_name, raw_value in dict(line.get("tempos_operacao", {}) or {}).items()
                if str(self.desktop_main.normalize_operacao_nome(op_name) or op_name or "").strip() and self._parse_float(raw_value, 0) > 0
            }
            if detailed_op_times:
                for op_name, unit_time in detailed_op_times.items():
                    if op_name not in planning_ops:
                        continue
                    total_time = unit_time * max(qtd_line, 0.0)
                    tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + total_time, 2)
                    if op_name == "Corte Laser":
                        esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + total_time, 2)
                        if not str(maquinas_operacao.get(op_name, "") or "").strip():
                            maquinas_operacao[op_name] = self.workcenter_default_resource(op_name, preferred=enc.get("posto_trabalho", ""))
            elif len(planning_ops) == 1:
                op_name = planning_ops[0]
                tempos_operacao[op_name] = round(float(tempos_operacao.get(op_name, 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if op_name == "Corte Laser":
                    esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if not str(maquinas_operacao.get(op_name, "") or "").strip():
                    maquinas_operacao[op_name] = self.workcenter_default_resource(op_name, preferred=enc.get("posto_trabalho", ""))
            elif "Corte Laser" in planning_ops:
                tempos_operacao["Corte Laser"] = round(float(tempos_operacao.get("Corte Laser", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                esp_bucket["tempo_min"] = round(float(esp_bucket.get("tempo_min", 0) or 0) + (tempo_peca * max(qtd_line, 0.0)), 2)
                if not str(maquinas_operacao.get("Corte Laser", "") or "").strip():
                    maquinas_operacao["Corte Laser"] = self.workcenter_default_resource("Corte Laser", preferred=enc.get("posto_trabalho", ""))
            raw_ref_interna = str(line.get("ref_interna", "") or "").strip()
            if raw_ref_interna and raw_ref_interna not in used_refs:
                ref_interna = raw_ref_interna
            else:
                ref_interna = str(self.desktop_main.next_ref_interna_unique(data, cliente_code, list(used_refs)))
            used_refs.add(ref_interna)
            ops_txt = self.quote_format_operacoes(line.get("operacao", ""))
            peca = {
                "id": f"PEC{piece_idx:05d}",
                "ref_interna": ref_interna,
                "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                "material": material,
                "espessura": espessura,
                "quantidade_pedida": qtd_line,
                "Operacoes": ops_txt,
                "Observacoes": str(line.get("descricao", "") or "").strip(),
                "desenho": str(line.get("desenho", "") or "").strip(),
                "tempo_peca_min": tempo_peca,
                "tempos_operacao": dict(line.get("tempos_operacao", {}) or {}),
                "custos_operacao": dict(line.get("custos_operacao", {}) or {}),
                "operacoes_detalhe": [dict(item or {}) for item in list(line.get("operacoes_detalhe", []) or []) if isinstance(item, dict)],
                "of": self.desktop_main.next_of_numero(data),
                "opp": self.desktop_main.next_opp_numero(data),
                "estado": "Preparacao",
                "produzido_ok": 0.0,
                "produzido_nok": 0.0,
                "inicio_producao": "",
                "fim_producao": "",
            }
            peca["operacoes_fluxo"] = self.desktop_main.build_operacoes_fluxo(ops_txt)
            piece_idx += 1
            mats[material]["espessuras"][espessura]["pecas"].append(peca)
            self.desktop_main.update_refs(data, peca["ref_interna"], peca["ref_externa"])
        enc["materiais"] = []
        for row in mats.values():
            row["espessuras"] = list(row["espessuras"].values())
            enc["materiais"].append(row)
        enc["montagem_itens"] = montagem_items
        enc["tempo_estimado"] = round(total_time, 2)
        enc["tempo"] = round(total_time / 60.0, 2) if total_time > 0 else 0.0
        data.setdefault("encomendas", []).append(enc)
        self._ensure_unique_order_piece_refs(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        orc["numero_encomenda"] = enc["numero"]
        if note:
            orc["nota_cliente"] = note
        orc["estado"] = "Convertido em Encomenda"
        self._save(force=True)
        return {
            "orcamento": self.orc_detail(numero),
            "encomenda": self.order_detail(enc["numero"]),
        }

    def _quote_purchase_need_key(self, kind: str, line: dict[str, Any]) -> str:
        if kind == "product":
            code = str(line.get("produto_codigo", "") or "").strip()
            if code:
                return f"product:{code}"
            return "product:new:" + self.desktop_main.norm_text(str(line.get("descricao", "") or "").strip())
        ref = str(line.get("stock_material_id", "") or "").strip()
        if ref:
            return f"material:{ref}"
        parts = [
            str(line.get("material", "") or "").strip(),
            str(line.get("espessura", "") or "").strip(),
            str(line.get("material_subtype", "") or line.get("calc_mode", "") or "").strip(),
            str(line.get("descricao", "") or "").strip(),
        ]
        return "material:new:" + "|".join(self.desktop_main.norm_text(part) for part in parts if part)

    def orc_purchase_needs(self, numero: str = "", lines: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        data = self.ensure_data()
        numero_txt = str(numero or "").strip()
        quote_lines = list(lines or [])
        if not quote_lines:
            orc = next((row for row in data.get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero_txt), None)
            if orc is None:
                raise ValueError("Orcamento nao encontrado.")
            quote_lines = list(orc.get("linhas", []) or [])
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(data.get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        grouped: dict[str, dict[str, Any]] = {}
        for raw in quote_lines:
            line = dict(raw or {})
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            if line_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                qty = max(0.0, self._parse_float(line.get("qtd", 0), 0))
                if qty <= 1e-9:
                    continue
                code = str(line.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                available = max(0.0, self._parse_float((product or {}).get("qty", 0), 0)) if product is not None else 0.0
                missing = qty if product is None else max(0.0, qty - available)
                if missing <= 1e-9:
                    continue
                key = self._quote_purchase_need_key("product", line)
                entry = grouped.setdefault(
                    key,
                    {
                        "kind": "product",
                        "ref": code,
                        "descricao": str(line.get("descricao", "") or (product or {}).get("descricao", "") or "").strip(),
                        "unid": str(line.get("produto_unid", "") or (product or {}).get("unid", "") or "UN").strip() or "UN",
                        "qtd": 0.0,
                        "qtd_disponivel": available,
                        "preco": round(self._parse_float((product or {}).get("p_compra", line.get("preco_unit", 0)), 0), 4),
                        "_product_pending_create": product is None or bool(line.get("_product_pending_create", False)),
                    },
                )
                entry["qtd"] = round(self._parse_float(entry.get("qtd", 0), 0) + missing, 2)
                continue
            if line_type != self.desktop_main.ORC_LINE_TYPE_PIECE or not self._quote_line_is_raw_material(line):
                continue
            qty = max(0.0, self._parse_float(line.get("qtd", 0), 0))
            if qty <= 1e-9:
                continue
            stock_id = str(line.get("stock_material_id", "") or "").strip()
            material_record = self.material_by_id(stock_id) if stock_id else None
            available = 0.0
            if isinstance(material_record, dict):
                available = max(
                    0.0,
                    self._parse_float(material_record.get("quantidade", 0), 0)
                    - self._parse_float(material_record.get("reservado", 0), 0),
                )
            missing = qty if material_record is None else max(0.0, qty - available)
            if missing <= 1e-9:
                continue
            formato = str(line.get("material_subtype", "") or line.get("calc_mode", "") or (material_record or {}).get("formato", "") or "Chapa").strip()
            if formato == "Stock MP":
                formato = str((material_record or {}).get("formato", "") or self.desktop_main.detect_materia_formato(material_record or {}) or "Chapa").strip()
            price = self._parse_float(line.get("price_base_value", 0), 0)
            if price <= 0 and isinstance(material_record, dict):
                price = self._parse_float(material_record.get("p_compra", material_record.get("preco_unid", 0)), 0)
            key = self._quote_purchase_need_key("material", line)
            entry = grouped.setdefault(
                key,
                {
                    "kind": "material",
                    "ref": stock_id,
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "unid": "UN",
                    "qtd": 0.0,
                    "qtd_disponivel": available,
                    "preco": round(price, 4),
                    "material": str(line.get("material", "") or (material_record or {}).get("material", "") or "").strip(),
                    "espessura": str(line.get("espessura", "") or (material_record or {}).get("espessura", "") or "").strip(),
                    "formato": formato or "Chapa",
                    "comprimento": round(self._parse_float(line.get("length_mm", (material_record or {}).get("comprimento", 0)), 0), 3),
                    "largura": round(self._parse_float(line.get("width_mm", (material_record or {}).get("largura", 0)), 0), 3),
                    "diametro": round(self._parse_float(line.get("diameter_mm", (material_record or {}).get("diametro", 0)), 0), 3),
                    "metros": round(self._parse_float(line.get("meters_per_unit", (material_record or {}).get("metros", 0)), 0), 4),
                    "kg_m": round(self._parse_float(line.get("kg_per_m", (material_record or {}).get("kg_m", 0)), 0), 4),
                    "peso_unid": round(self._parse_float(line.get("stock_metric_value", (material_record or {}).get("peso_unid", 0)), 0), 4),
                    "_material_pending_create": material_record is None,
                    "_material_manual": material_record is None,
                },
            )
            entry["qtd"] = round(self._parse_float(entry.get("qtd", 0), 0) + missing, 2)
        rows = [row for row in grouped.values() if self._parse_float(row.get("qtd", 0), 0) > 0]
        rows.sort(key=lambda row: (str(row.get("kind", "")), str(row.get("ref", "") or row.get("descricao", ""))))
        return rows

    def orc_create_purchase_quote(self, numero: str, lines: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        needs = self.orc_purchase_needs(numero_txt, lines)
        if not needs:
            raise ValueError("Nao existem necessidades de compra nas linhas do orcamento.")
        note = self.ne_save(
            {
                "fornecedor": "",
                "fornecedor_id": "",
                "contacto": "",
                "obs": f"Pedido de cotacao gerado a partir do orcamento {numero_txt}".strip(),
                "lines": [
                    (
                        {
                            "ref": str(need.get("ref", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip() or str(need.get("material", "") or "").strip(),
                            "origem": "Materia-prima",
                            "qtd": round(self._parse_float(need.get("qtd", 0), 0), 2),
                            "unid": str(need.get("unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "material": str(need.get("material", "") or "").strip(),
                            "espessura": str(need.get("espessura", "") or "").strip(),
                            "formato": str(need.get("formato", "") or "Chapa").strip() or "Chapa",
                            "comprimento": self._parse_float(need.get("comprimento", 0), 0),
                            "largura": self._parse_float(need.get("largura", 0), 0),
                            "diametro": self._parse_float(need.get("diametro", 0), 0),
                            "metros": self._parse_float(need.get("metros", 0), 0),
                            "kg_m": self._parse_float(need.get("kg_m", 0), 0),
                            "peso_unid": self._parse_float(need.get("peso_unid", 0), 0),
                            "_material_pending_create": bool(need.get("_material_pending_create", False)),
                            "_material_manual": bool(need.get("_material_manual", False)),
                        }
                        if str(need.get("kind", "") or "") == "material"
                        else {
                            "ref": str(need.get("ref", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip(),
                            "origem": "Produto",
                            "qtd": round(self._parse_float(need.get("qtd", 0), 0), 2),
                            "unid": str(need.get("unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "_product_pending_create": bool(need.get("_product_pending_create", False)),
                        }
                    )
                    for need in needs
                ],
            }
        )
        note_number = str(note.get("numero", "") or "").strip()
        return {"numero": note_number, "line_count": len(list(note.get("linhas", []) or [])), "needs": needs, "detail": self.ne_detail(note_number)}

    def orc_suggest_notes(self, payload: dict[str, Any]) -> str:
        helper = self._orc_render_helper()
        lines = self.orc_actions._build_orc_notes_lines(helper, payload)
        return "\n".join([str(line or "").strip() for line in lines if str(line or "").strip()])

