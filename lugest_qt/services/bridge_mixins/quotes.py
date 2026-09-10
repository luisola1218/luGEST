from __future__ import annotations
from lugest_qt.services.quote_routing_composition import line_routing
from lugest_qt.services.quote_purchase_composition import purchase_needs, purchase_quote
from lugest_qt.services.quote_order_composition import quote_conversion
from lugest_modules.quotes.application.assemblies import normalize_item, price_item, refresh_model, expand_model, technical_sheet
from lugest_qt.services.assembly_composition import assembly_rules, assembly_refresh, assembly_catalog, assembly_queries
from lugest_modules.quotes.application.assembly_refresh import assign_parameter_codes
from lugest_qt.services.quote_queries_composition import quote_queries
from lugest_qt.services.quote_commands_composition import quote_commands
from lugest_modules.quotes.application.line_normalization import normalize_line
from lugest_qt.services.quote_rules_composition import normalize_line_ports
from lugest_modules.quotes.infrastructure.assembly_report import render_assembly_sheet
from lugest_qt.services.quote_rules_composition import render_assembly_sheet_ports
from lugest_modules.quotes.infrastructure.nesting_report import render_nesting_study
from lugest_qt.services.quote_rules_composition import render_nesting_study_ports
from lugest_qt.services.quote_nesting_composition import nesting_sql_store, nesting_study_service

import uuid

import copy
import json
import math
import os
import tempfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text


ORC_STANDARD_IVA_PERC = 23.0
ORC_DEFAULT_DELIVERY_TEXT = "A combinar com o departamento de planeamento."


class QuotesBridgeMixin:
    """Quote, assembly model, and quote-to-order operations for the Qt bridge."""

    def _quote_standard_iva_perc(self) -> float:
        return ORC_STANDARD_IVA_PERC

    def _quote_default_delivery_text(self) -> str:
        return ORC_DEFAULT_DELIVERY_TEXT

    def _normalize_quote_discount_mode(self, value: Any) -> str:
        mode = str(value or "").strip().lower()
        return mode if mode in {"total", "lotes_espessura"} else "total"

    def _normalize_quote_discount_groups(self, value: Any) -> list[str]:
        groups: list[str] = []
        for item in list(value or []):
            clean = str(item or "").strip()
            if clean and clean not in groups:
                groups.append(clean)
        return groups

    def _peek_next_orc_number(self) -> str:
        data = self.ensure_data()
        try:
            return str(self.desktop_main.peek_next_orc_numero(data))
        except Exception:
            try:
                seq = int(data.get("orc_seq", 1) or 1)
            except Exception:
                seq = 1
            year = int(getattr(datetime.now(), "year", 0) or 0)
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
        return quote_queries(self).rows(filter_text, state_filter, year)

    def orc_available_years(self) -> list[str]:
        return quote_queries(self).years()

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
        return nesting_sql_store(self).ensure_table(conn)

    def _mysql_orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        return nesting_sql_store(self).studies(numero)

    def _mysql_save_orc_nesting_study(self, numero: str, group_key: str, group_label: str, payload: dict[str, Any]) -> None:
        return nesting_sql_store(self).save(numero, group_key, group_label, payload)

    def _mysql_delete_orc_nesting_studies(self, numero: str, group_key: str = "") -> None:
        return nesting_sql_store(self).delete(numero, group_key)

    def orc_nesting_studies(self, numero: str) -> dict[str, Any]:
        return nesting_study_service(self).studies(numero)

    def orc_save_nesting_study(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        return nesting_study_service(self).save(numero, payload)

    def orc_detail(self, numero: str) -> dict[str, Any]:
        return quote_queries(self).detail(numero)

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
        for row in list(self.ensure_data().get("conjuntos_modelo", []) or []) + list(self.ensure_data().get("conjuntos", []) or []):
            codigo = str((row or {}).get("codigo", "") or "").strip().upper()
            digits = "".join(ch for ch in codigo if ch.isdigit())
            if digits:
                try:
                    highest = max(highest, int(digits))
                except Exception:
                    continue
        return f"MOD{highest + 1:04d}"

    def conjunto_next_param_codigo(self) -> str:
        highest = 0
        missing = 0
        for row in list(self.ensure_data().get("conjuntos", []) or []):
            raw = str((row or {}).get("param_codigo", "") or "").strip()
            digits = "".join(ch for ch in raw if ch.isdigit())
            if digits:
                highest = max(highest, int(digits))
            else:
                missing += 1
        return f"{highest + missing + 1:04d}"

    def _ensure_conjunto_param_codes(self) -> bool:
        return assign_parameter_codes(list(self.ensure_data().get("conjuntos", []) or []))

    def _conjunto_find_quote_source(self, item: dict[str, Any], conjunto_codigo: str) -> tuple[dict[str, Any] | None, str]:
        ref = str(item.get("source_ref_externa", "") or item.get("ref_externa", "") or "").strip()
        quote_number = str(item.get("source_quote_number", "") or "").strip()
        operation_norm = self.desktop_main.norm_text(str(item.get("operacao", "") or ""))
        if not ref or ("laser" not in operation_norm and not str(item.get("desenho", "") or "").strip()):
            return None, quote_number
        quotes = list(self.ensure_data().get("orcamentos", []) or [])
        if quote_number:
            quotes = sorted(quotes, key=lambda row: str(row.get("numero", "") or "") != quote_number)
        else:
            quotes = list(reversed(quotes))
        fallback: tuple[dict[str, Any] | None, str] = (None, "")
        for quote in quotes:
            number = str(quote.get("numero", "") or "").strip()
            for line in list(quote.get("linhas", []) or []):
                if str(line.get("ref_externa", "") or "").strip() != ref:
                    continue
                if str(line.get("conjunto_codigo", "") or "").strip() == conjunto_codigo:
                    return line, number
                if fallback[0] is None:
                    fallback = (line, number)
            if quote_number and number == quote_number and fallback[0] is not None:
                return fallback
        return fallback

    def _conjunto_live_item(self, raw_item: dict[str, Any], conjunto_codigo: str) -> tuple[dict[str, Any], bool]:
        return price_item(assembly_rules(self), raw_item, conjunto_codigo)

    def _conjunto_refresh_model_prices(self, model: dict[str, Any]) -> bool:
        refreshed, changed = refresh_model(assembly_rules(self), model)
        model.update(refreshed)
        return changed

    def conjunto_refresh_prices(self, codigo: str = "") -> dict[str, Any]:
        return assembly_refresh(self).refresh(codigo)

    def _normalize_assembly_model_item(self, payload: dict[str, Any]) -> dict[str, Any]:
        return normalize_item(assembly_rules(self), payload)

    def assembly_model_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        return assembly_queries(self, templates=True).template_rows(filter_text)

    def assembly_model_detail(self, codigo: str) -> dict[str, Any]:
        return assembly_queries(self, templates=True).template_detail(codigo)

    @staticmethod
    def _normalize_conjunto_technical_sheet(raw: Any) -> dict[str, str]:
        return technical_sheet(raw)

    def assembly_model_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.assembly_model_detail(assembly_catalog(self, templates=True).save(payload))

    def assembly_model_remove(self, codigo: str) -> None:
        return assembly_catalog(self, templates=True).remove(codigo)

    def assembly_model_expand(self, codigo: str, quantity: Any = 1) -> list[dict[str, Any]]:
        return expand_model(assembly_rules(self), self.assembly_model_detail(codigo), quantity, uuid.uuid4().hex[:12].upper())

    def conjunto_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        self.conjunto_refresh_prices()
        return assembly_queries(self).rows(filter_text)

    def conjunto_detail(self, codigo: str) -> dict[str, Any]:
        self.conjunto_refresh_prices(codigo)
        return assembly_queries(self).detail(codigo)

    def conjunto_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.conjunto_detail(assembly_catalog(self).save(payload))

    def conjunto_remove(self, codigo: str) -> None:
        return assembly_catalog(self).remove(codigo)

    def conjunto_expand(self, codigo: str, quantity: Any = 1) -> list[dict[str, Any]]:
        return expand_model(assembly_rules(self), self.conjunto_detail(codigo), quantity, uuid.uuid4().hex[:12].upper())

    def _normalize_orc_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        return normalize_line(normalize_line_ports(self), payload)

    def orc_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.orc_detail(quote_commands(self).save(payload))

    def orc_remove(self, numero: str) -> None:
        return quote_commands(self).remove(numero)

    def orc_set_state(self, numero: str, estado: str) -> dict[str, Any]:
        return self.orc_detail(quote_commands(self).set_state(numero, estado))

    def _orc_render_helper(self) -> Any:
        helper = SimpleNamespace(data=self.ensure_data())
        helper._extract_orc_operacoes = lambda orc=None: self.orc_actions._extract_orc_operacoes(helper, orc)
        helper._build_orc_notes_lines = lambda orc: self.orc_actions._build_orc_notes_lines(helper, orc)
        return helper

    def orc_render_pdf(self, numero: str, path: str | Path) -> Path:
        numero = str(numero or "").strip()
        orc = next((row for row in self.ensure_data().get("orcamentos", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if orc is None:
            raise ValueError("Orçamento não encontrado.")
        target = Path(path)
        helper = self._orc_render_helper()
        self.orc_actions.render_orc_pdf(helper, str(target), orc)
        return target

    def orc_open_pdf(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_orcamento_{str(numero or '').strip()}.pdf"
        self.orc_render_pdf(numero, target)
        os.startfile(str(target))
        return target

    def orc_print_pdf(self, numero: str) -> Path:
        target = Path(tempfile.gettempdir()) / f"lugest_orcamento_{str(numero or '').strip()}_print.pdf"
        self.orc_render_pdf(numero, target)
        try:
            os.startfile(str(target), "print")
        except Exception:
            os.startfile(str(target))
        return target

    def conjunto_sheet_pdf(self, codigo: str, output_path: str | Path | None = None) -> Path:
        return render_assembly_sheet(render_assembly_sheet_ports(self), codigo, output_path)


    def conjunto_open_sheet_pdf(self, codigo: str) -> Path:
        target = self.conjunto_sheet_pdf(codigo)
        os.startfile(str(target))
        return target

    def orc_render_nesting_study_pdf(self, numero: str, path: str | Path, group_key: str = "") -> Path:
        return render_nesting_study(render_nesting_study_ports(self), numero, path, group_key)

    def orc_open_nesting_study_pdf(self, numero: str, group_key: str = "") -> Path:
        safe_group = "".join(ch if ch.isalnum() else "_" for ch in str(group_key or "").strip()) or "grupo"
        target = Path(tempfile.gettempdir()) / f"lugest_nesting_{str(numero or '').strip()}_{safe_group}.pdf"
        self.orc_render_nesting_study_pdf(numero, target, group_key=group_key)
        os.startfile(str(target))
        return target

    def _quote_line_operations_value(self, line: dict[str, Any] | None = None) -> list[str]:
        return line_routing(self).operations(line)

    def _quote_line_operations_text(self, line: dict[str, Any] | None = None) -> str:
        return line_routing(self).operations_text(line)

    def _quote_line_is_production_ready(self, line: dict[str, Any] | None = None) -> bool:
        return line_routing(self).production_ready(line)

    def _quote_line_is_raw_material(self, line: dict[str, Any] | None = None) -> bool:
        return line_routing(self).raw_material(line)

    def _quote_line_looks_stock_material_ref(self, value: Any) -> bool:
        return line_routing(self).stock_material_reference(value)

    def _quote_line_production_route(self, line: dict[str, Any] | None = None) -> str:
        return line_routing(self).production_route(line)

    def orc_convert_to_order(self, numero: str, nota_cliente: str = "") -> dict[str, Any]:
        order_number = quote_conversion(self).convert(numero, nota_cliente)
        return {"orcamento": self.orc_detail(numero), "encomenda": self.order_detail(order_number)}

    def _quote_purchase_need_key(self, kind: str, line: dict[str, Any]) -> str:
        return purchase_needs(self).key(kind, line)

    def orc_purchase_needs(self, numero: str = "", lines: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
        return purchase_needs(self).rows(numero, lines)

    def orc_create_purchase_quote(self, numero: str, lines: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        return purchase_quote(self).create(numero, lines)

    def orc_suggest_notes(self, payload: dict[str, Any]) -> str:
        helper = self._orc_render_helper()
        lines = self.orc_actions._build_orc_notes_lines(helper, payload)
        return "\n".join([str(line or "").strip() for line in lines if str(line or "").strip()])

