from __future__ import annotations

import math
import os
import re
import shutil
import tempfile
import time
from lugest_infra.pdf.text import (
    clip_text as _pdf_clip_text,
    fit_font_size as _pdf_fit_font_size,
    wrap_text as _pdf_wrap_text,
)
from pathlib import Path
from typing import Any


class OrderReportsBackendMixin:
    """Legacy adapter for order reports; see BACKEND_GUIDE.md."""

    def order_fabrication_pdf(
        self,
        numero: str,
        output_path: str | Path | None = None,
        selected_groups: list[dict[str, Any]] | None = None,
    ) -> Path:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        changed = self._ensure_order_fabrication_order(enc)
        digest_cache: dict[str, str] = {}
        for piece in self.desktop_main.encomenda_pecas(enc):
            if self._apply_piece_pdf_references(piece, self._piece_pdf_references(piece, digest_cache=digest_cache), digest_cache=digest_cache):
                changed = True
        if self._sync_order_piece_documents_from_quote(enc):
            changed = True
        if self._sync_order_piece_metrics_from_quote(enc):
            changed = True
        if changed:
            self._save(force=True)
        detail = self.order_detail(numero)
        pieces = list(detail.get("pieces", []) or [])
        selected_lookup: set[tuple[str, str]] = set()
        for group in list(selected_groups or []):
            if not isinstance(group, dict):
                continue
            material_txt = str(group.get("material", "") or "").strip()
            esp_txt = str(group.get("espessura", "") or "").strip()
            if material_txt and esp_txt:
                selected_lookup.add((material_txt, esp_txt))
        if selected_lookup:
            pieces = [
                piece
                for piece in pieces
                if (
                    str(piece.get("material", "") or "").strip(),
                    str(piece.get("espessura", "") or "").strip(),
                )
                in selected_lookup
            ]
        if not pieces:
            raise ValueError("A ordem de fabrico não tem peças para as espessuras selecionadas.")
        of_code = str(detail.get("of_codigo", "") or self._order_of_code(enc)).strip()
        if output_path:
            target = Path(output_path)
        elif selected_lookup:
            suffix = "_".join(
                re.sub(r"[^A-Za-z0-9]+", "-", f"{material}-{esp}").strip("-")
                for material, esp in sorted(selected_lookup)
            )[:80]
            target = Path(tempfile.gettempdir()) / f"lugest_ordem_fabrico_{of_code}_{suffix or 'parcial'}.pdf"
        else:
            target = Path(tempfile.gettempdir()) / f"lugest_ordem_fabrico_{of_code}.pdf"
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except ModuleNotFoundError:
            lines = [
                "Ordem de Fabrico",
                f"OF: {of_code}",
                f"Encomenda: {detail.get('numero', '-')}",
                f"Cliente: {detail.get('cliente', '-') or '-'} - {detail.get('cliente_nome', '') or ''}".strip(" -"),
                f"Data: {str((detail.get('ordem_fabrico') or {}).get('data', '') or '')[:10]}",
                "",
            ]
            fallback_groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
            for piece in pieces:
                key = (str(piece.get("material", "") or "-").strip() or "-", str(piece.get("espessura", "") or "-").strip() or "-")
                fallback_groups.setdefault(key, []).append(piece)
            for (material_txt, esp_txt), group_rows in sorted(fallback_groups.items(), key=lambda item: (item[0][0].lower(), item[0][1].lower())):
                lines.append(f"Espessura: {material_txt} | {esp_txt} mm | {len(group_rows)} peca(s)")
                lines.append(f"Codigo grupo: GRP|{of_code}|{material_txt}|{esp_txt}")
                for piece in group_rows:
                    lines.append(
                        f"- {piece.get('ref_interna', '-') or '-'} | {piece.get('ref_externa', '-') or '-'} | "
                        f"Qtd {piece.get('qtd_plan', '0')} | Ops: {piece.get('operacoes', '-') or '-'}"
                    )
                lines.append("")
            self._write_basic_pdf(target, lines)
            return target
        page_w, page_h = A4
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        margin = 12 * mm
        inner_w = page_w - (margin * 2)
        header_h = 38 * mm
        table_header_h = 8 * mm
        row_h = 17.0 * mm
        group_h = 10.5 * mm
        footer_h = 9 * mm
        columns = [
            ("Ref.", 27 * mm),
            ("Ref. externa", 50 * mm),
            ("Material", 26 * mm),
            ("Esp.", 12 * mm),
            ("Qtd", 12 * mm),
            ("Operacoes / OPP", inner_w - (27 + 50 + 26 + 12 + 12) * mm),
        ]
        rows_per_page = max(1, int((page_h - (margin * 2) - header_h - table_header_h - footer_h) // row_h))
        grouped_pieces: list[tuple[str, str, list[dict[str, Any]]]] = []
        groups_map: dict[tuple[str, str], list[dict[str, Any]]] = {}
        for piece in pieces:
            key = (str(piece.get("material", "") or "-").strip() or "-", str(piece.get("espessura", "") or "-").strip() or "-")
            groups_map.setdefault(key, []).append(piece)
        for key, rows in sorted(groups_map.items(), key=lambda item: (item[0][0].lower(), item[0][1].lower())):
            grouped_pieces.append((key[0], key[1], rows))
        def _piece_pdf_docs(piece: dict[str, Any]) -> list[Path]:
            docs: list[Path] = []
            for raw_doc in self._piece_pdf_references(piece, digest_cache=digest_cache):
                doc_txt = str(raw_doc or "").strip()
                if not doc_txt or not doc_txt.lower().endswith(".pdf"):
                    continue
                doc_path = self._resolve_file_reference(doc_txt)
                if doc_path is None:
                    doc_path = Path(doc_txt)
                if doc_path.exists() and doc_path not in docs:
                    docs.append(doc_path)
            return docs

        technical_groups: list[dict[str, Any]] = []
        for material_group, esp_group, group_rows in grouped_pieces:
            docs_rows: list[dict[str, Any]] = []
            for piece in group_rows:
                docs = _piece_pdf_docs(piece)
                if docs:
                    docs_rows.append({"piece": piece, "docs": docs})
            if docs_rows:
                technical_groups.append({"material": material_group, "espessura": esp_group, "pieces": group_rows, "docs_rows": docs_rows})
        reservation_lines: list[str] = []
        for reserva in list(enc.get("reservas", []) or []):
            if not isinstance(reserva, dict):
                continue
            mat_txt = str(reserva.get("material", "") or "-").strip() or "-"
            esp_txt = str(reserva.get("espessura", "") or "-").strip() or "-"
            if selected_lookup and (mat_txt, esp_txt) not in selected_lookup:
                continue
            qty_txt = str(reserva.get("quantidade", reserva.get("qtd", reserva.get("chapas", ""))) or "-").strip() or "-"
            state_txt = str(reserva.get("estado", reserva.get("status", reserva.get("material_estado", ""))) or "").strip() or "Reservado"
            reservation_lines.append(f"{mat_txt} | {esp_txt} mm | {qty_txt} chapa(s) | {state_txt}")
        product_sheets = [
            dict(row or {})
            for row in list(detail.get("produto_fichas", []) or [])
            if isinstance(row, dict) and str(row.get("codigo", "") or "").strip()
        ]
        known_sheet_codes = {str(row.get("codigo", "") or "").strip() for row in product_sheets}
        legacy_sheet_groups: dict[str, set[str]] = {}
        for source_row in [*pieces, *list(detail.get("montagem_items", []) or [])]:
            sheet_code = str(source_row.get("conjunto_codigo", "") or "").strip()
            if not sheet_code or sheet_code in known_sheet_codes:
                continue
            group_key = str(source_row.get("grupo_uuid", "") or "").strip() or sheet_code
            legacy_sheet_groups.setdefault(sheet_code, set()).add(group_key)
        for sheet_code, group_keys in legacy_sheet_groups.items():
            stored_sheet: dict[str, Any] = {}
            for detail_fn in (self.conjunto_detail, self.assembly_model_detail):
                try:
                    stored_sheet = dict(detail_fn(sheet_code) or {})
                except Exception:
                    stored_sheet = {}
                if stored_sheet:
                    break
            product_sheets.append(
                {
                    "codigo": sheet_code,
                    "param_codigo": str(stored_sheet.get("param_codigo", "") or "").strip(),
                    "descricao": str(stored_sheet.get("descricao", "") or sheet_code).strip(),
                    "notas": str(stored_sheet.get("notas", "") or "").strip(),
                    "ficha_tecnica": dict(stored_sheet.get("ficha_tecnica", {}) or {}),
                    "quantidade_conjuntos": max(1, len(group_keys)),
                }
            )
        visual_units = len(pieces) + (len(grouped_pieces) * max(1, math.ceil(group_h / row_h)))
        total_pages = max(1, math.ceil(max(1, visual_units) / rows_per_page)) + len(product_sheets)
        montagem_stock_items = [
            row
            for row in list(detail.get("montagem_items", []) or [])
            if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) == self.desktop_main.ORC_LINE_TYPE_PRODUCT
            or bool(str(row.get("stock_material_id", "") or "").strip())
        ]
        montagem_stock_code = f"COMP|{of_code}" if montagem_stock_items else ""
        montagem_component_groups: list[tuple[str, list[dict[str, Any]]]] = []
        if montagem_stock_items:
            grouped_components: dict[str, list[dict[str, Any]]] = {}
            for item_index, item in enumerate(montagem_stock_items):
                esp_txt = str(item.get("espessura", "") or "").strip() or "Sem espessura"
                code_txt = str(item.get("produto_codigo", "") or item.get("stock_material_id", "") or self._montagem_item_key(item) or "-").strip() or "-"
                qty_txt = self._fmt(item.get("qtd_planeada", 0))
                unidade_txt = str(item.get("produto_unid", "") or "UN").strip() or "UN"
                grouped_components.setdefault(esp_txt, []).append(
                    {
                        "barcode": f"CPI|{of_code}|{item_index}",
                        "codigo": code_txt,
                        "descricao": str(item.get("descricao", "") or item.get("material", "") or "-").strip() or "-",
                        "quantidade": f"{qty_txt} {unidade_txt}".strip(),
                    }
                )
            montagem_component_groups = sorted(
                grouped_components.items(),
                key=lambda item: (
                    self._parse_float(item[0].replace("Sem espessura", "0"), 0),
                    item[0].lower(),
                ),
            )
        cover_target = target
        if technical_groups or montagem_component_groups:
            cover_target = Path(tempfile.gettempdir()) / f"lugest_of_cover_{of_code}_{int(time.time() * 1000)}.pdf"
        canvas_obj = pdf_canvas.Canvas(str(cover_target), pagesize=A4)

        def draw_header(page_number: int) -> float:
            top_y = page_h - margin
            canvas_obj.setFillColor(colors.white)
            canvas_obj.rect(0, 0, page_w, page_h, stroke=0, fill=1)
            canvas_obj.setFillColor(palette["primary"])
            canvas_obj.rect(0, page_h - 3 * mm, page_w, 3 * mm, stroke=0, fill=1)
            canvas_obj.setFillColor(colors.HexColor("#FFFFFF"))
            canvas_obj.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas_obj.roundRect(margin, top_y - header_h, inner_w, header_h, 5, stroke=1, fill=1)
            self._draw_operator_logo_plate(canvas_obj, palette, logo_path, margin + 8, top_y - 23 * mm, 34 * mm, 15 * mm, radius=4, padding_x=3, padding_y=2)
            canvas_obj.setFillColor(colors.HexColor("#020617"))
            canvas_obj.setFont("Helvetica-Bold", 16)
            canvas_obj.drawString(margin + 48 * mm, top_y - 10 * mm, self._operator_pdf_text("Ordem de Fabrico"))
            canvas_obj.setFont("Helvetica", 7.5)
            client_line = f"{detail.get('cliente', '-') or '-'} - {detail.get('cliente_nome', '') or ''}".strip(" -")
            header_meta = [
                f"Encomenda: {detail.get('numero', '-')}",
                f"Cliente: {client_line or '-'}",
                f"Data OF: {str((detail.get('ordem_fabrico') or {}).get('data', '') or '')[:10]}",
                f"Pecas: {len(pieces)}",
            ]
            meta_y = top_y - 16 * mm
            for line in header_meta:
                canvas_obj.drawString(margin + 48 * mm, meta_y, self._operator_pdf_text(_pdf_clip_text(line, 80 * mm, "Helvetica", 7.5)))
                meta_y -= 4 * mm
            barcode_x = page_w - margin - 66 * mm
            canvas_obj.setFont("Helvetica-Bold", 9)
            canvas_obj.drawCentredString(barcode_x + 33 * mm, top_y - 8 * mm, self._operator_pdf_text(of_code))
            self._draw_code128_fit(canvas_obj, of_code, barcode_x, top_y - 23 * mm, 66 * mm, 13 * mm, min_bar_width=0.36, max_bar_width=0.82)
            canvas_obj.setFont("Helvetica", 7)
            canvas_obj.setFillColor(colors.HexColor("#64748B"))
            canvas_obj.drawRightString(page_w - margin - 8, top_y - 30 * mm, self._operator_pdf_text(f"Pagina {page_number}/{total_pages}"))
            cursor_y = top_y - header_h - 4 * mm
            if page_number == 1:
                box_h = (8 + (max(1, min(4, len(reservation_lines))) * 4.2)) * mm
                canvas_obj.setFillColor(colors.HexColor("#FFF8EB"))
                canvas_obj.setStrokeColor(colors.HexColor("#E4C37F"))
                canvas_obj.roundRect(margin, cursor_y - box_h, inner_w, box_h, 4, stroke=1, fill=1)
                canvas_obj.setFillColor(colors.HexColor("#7A3E00"))
                canvas_obj.setFont("Helvetica-Bold", 7.2)
                canvas_obj.drawString(margin + 3 * mm, cursor_y - 4.6 * mm, self._operator_pdf_text("Chapas cativadas"))
                canvas_obj.setFont("Helvetica", 6.7)
                if reservation_lines:
                    line_y = cursor_y - 8.6 * mm
                    for line in reservation_lines[:4]:
                        canvas_obj.drawString(margin + 3 * mm, line_y, self._operator_pdf_text(_pdf_clip_text(line, inner_w - 8 * mm, "Helvetica", 6.7)))
                        line_y -= 4.2 * mm
                    if len(reservation_lines) > 4:
                        canvas_obj.drawRightString(page_w - margin - 3 * mm, cursor_y - 8.6 * mm, self._operator_pdf_text(f"+{len(reservation_lines) - 4} reserva(s)"))
                else:
                    canvas_obj.drawString(margin + 3 * mm, cursor_y - 8.6 * mm, self._operator_pdf_text("Sem chapas cativadas registadas."))
                cursor_y -= box_h + 3 * mm
            return cursor_y

        def draw_table_header(y_pos: float) -> float:
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.rect(margin, y_pos - table_header_h, inner_w, table_header_h, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["primary_dark"])
            canvas_obj.setFont("Helvetica-Bold", 6.8)
            x = margin
            for label, width in columns:
                canvas_obj.drawString(x + 2.2, y_pos - 5.2 * mm, self._operator_pdf_text(label))
                x += width
            return y_pos - table_header_h

        def draw_footer(page_number: int) -> None:
            canvas_obj.setFillColor(colors.HexColor("#64748B"))
            canvas_obj.setFont("Helvetica", 6.8)
            canvas_obj.drawString(margin, margin - 2, self._operator_pdf_text(f"Impresso em {printed_at}"))
            canvas_obj.drawRightString(page_w - margin, margin - 2, self._operator_pdf_text(f"LUGEST | OF {of_code} | {page_number}/{total_pages}"))

        def draw_product_sheet(snapshot: dict[str, Any], page_number: int) -> None:
            technical = dict(snapshot.get("ficha_tecnica", {}) or {})
            code_txt = str(snapshot.get("codigo", "") or "-").strip() or "-"
            param_txt = str(snapshot.get("param_codigo", "") or "-").strip() or "-"
            name_txt = str(snapshot.get("descricao", "") or code_txt).strip() or code_txt
            quantity_txt = self._fmt(snapshot.get("quantidade_conjuntos", 1))
            top_y = page_h - margin

            canvas_obj.setFillColor(colors.white)
            canvas_obj.rect(0, 0, page_w, page_h, stroke=0, fill=1)
            canvas_obj.setFillColor(palette["primary"])
            canvas_obj.rect(0, page_h - 3 * mm, page_w, 3 * mm, stroke=0, fill=1)

            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(colors.HexColor("#CBD5E1"))
            canvas_obj.roundRect(margin, top_y - 38 * mm, inner_w, 38 * mm, 5, stroke=1, fill=1)
            self._draw_operator_logo_plate(
                canvas_obj,
                palette,
                logo_path,
                margin + 8,
                top_y - 23 * mm,
                34 * mm,
                15 * mm,
                radius=4,
                padding_x=3,
                padding_y=2,
            )
            canvas_obj.setFillColor(colors.HexColor("#64748B"))
            canvas_obj.setFont("Helvetica-Bold", 7.2)
            canvas_obj.drawString(margin + 48 * mm, top_y - 7 * mm, self._operator_pdf_text("FICHA DE PRODUTO PARA PRODUCAO"))
            canvas_obj.setFillColor(colors.HexColor("#020617"))
            canvas_obj.setFont("Helvetica-Bold", 15)
            canvas_obj.drawString(
                margin + 48 * mm,
                top_y - 14 * mm,
                self._operator_pdf_text(_pdf_clip_text(name_txt, 91 * mm, "Helvetica-Bold", 15)),
            )
            canvas_obj.setFont("Helvetica", 7.5)
            canvas_obj.drawString(
                margin + 48 * mm,
                top_y - 20 * mm,
                self._operator_pdf_text(
                    _pdf_clip_text(
                        f"Codigo: {code_txt} | Param.: {param_txt} | Familia: {technical.get('familia_produto', '') or '-'} | Qtd: {quantity_txt}",
                        92 * mm,
                        "Helvetica",
                        7.5,
                    )
                ),
            )
            canvas_obj.drawString(
                margin + 48 * mm,
                top_y - 26 * mm,
                self._operator_pdf_text(
                    _pdf_clip_text(
                        f"Modelo/versao: {technical.get('modelo_versao', '') or '-'} | Aplicacao: {technical.get('aplicacao', '') or '-'}",
                        92 * mm,
                        "Helvetica",
                        7.5,
                    )
                ),
            )
            barcode_x = page_w - margin - 47 * mm
            self._draw_code128_fit(canvas_obj, param_txt if param_txt != "-" else code_txt, barcode_x, top_y - 25 * mm, 43 * mm, 12 * mm, min_bar_width=0.28, max_bar_width=0.62)
            canvas_obj.setFont("Helvetica-Bold", 5.6)
            canvas_obj.drawCentredString(barcode_x + 21.5 * mm, top_y - 28 * mm, self._operator_pdf_text(f"PARAM {param_txt}" if param_txt != "-" else code_txt))

            def field_box(x: float, y_top: float, width: float, label: str, value: str, height: float = 19 * mm) -> None:
                canvas_obj.setFillColor(colors.HexColor("#F8FAFC"))
                canvas_obj.setStrokeColor(colors.HexColor("#D7DEE8"))
                canvas_obj.roundRect(x, y_top - height, width, height, 4, stroke=1, fill=1)
                canvas_obj.setFillColor(colors.HexColor("#475569"))
                canvas_obj.setFont("Helvetica-Bold", 6.5)
                canvas_obj.drawString(x + 3 * mm, y_top - 5 * mm, self._operator_pdf_text(label.upper()))
                canvas_obj.setFillColor(colors.HexColor("#0F172A"))
                canvas_obj.setFont("Helvetica", 7.2)
                lines = _pdf_wrap_text(str(value or "-") or "-", "Helvetica", 7.2, width - 6 * mm, max_lines=4) or ["-"]
                line_y = y_top - 10 * mm
                for wrapped in lines:
                    canvas_obj.drawString(x + 3 * mm, line_y, self._operator_pdf_text(wrapped))
                    line_y -= 3.7 * mm

            content_top = top_y - 44 * mm
            half_w = (inner_w - 4 * mm) / 2
            field_box(margin, content_top, half_w, "Configuracao principal", str(technical.get("configuracao", "") or "-"), 18 * mm)
            field_box(margin + half_w + 4 * mm, content_top, half_w, "Dimensoes / capacidade", str(technical.get("dimensoes_gerais", "") or "-"), 18 * mm)
            content_top -= 22 * mm
            field_box(margin, content_top, half_w, "Materiais / acabamentos", str(technical.get("materiais_acabamentos", "") or "-"), 18 * mm)
            field_box(margin + half_w + 4 * mm, content_top, half_w, "Referencia de producao", f"OF {of_code} | Encomenda {detail.get('numero', '-')}", 18 * mm)
            content_top -= 24 * mm
            field_box(margin, content_top, inner_w, "Caracteristicas e desempenho", str(technical.get("caracteristicas", "") or "-"), 30 * mm)
            content_top -= 34 * mm
            field_box(margin, content_top, inner_w, "Requisitos de instalacao", str(technical.get("requisitos_instalacao", "") or "-"), 27 * mm)
            content_top -= 31 * mm
            field_box(margin, content_top, half_w, "Normas / conformidade", str(technical.get("normas_conformidade", "") or "-"), 28 * mm)
            field_box(margin + half_w + 4 * mm, content_top, half_w, "Controlo de qualidade", str(technical.get("controlo_qualidade", "") or "-"), 28 * mm)
            content_top -= 32 * mm
            notes_txt = str(snapshot.get("notas", "") or "").strip()
            field_box(margin, content_top, inner_w, "Observacoes de fabrico", notes_txt or "Sem observacoes adicionais.", 28 * mm)

            canvas_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
            canvas_obj.line(margin, margin + 4.5 * mm, page_w - margin, margin + 4.5 * mm)
            canvas_obj.setFillColor(colors.HexColor("#64748B"))
            canvas_obj.setFont("Helvetica", 6.2)
            canvas_obj.drawString(margin, margin + 1.4 * mm, self._operator_pdf_text(f"Snapshot tecnico da encomenda | Impresso em {printed_at}"))
            canvas_obj.drawRightString(page_w - margin, margin + 1.4 * mm, self._operator_pdf_text(f"LUGEST | OF {of_code} | {page_number}/{total_pages}"))

        page_number = 1
        y = draw_table_header(draw_header(page_number))
        row_counter = 0

        def ensure_space(height: float) -> None:
            nonlocal page_number, y, row_counter
            if y - height >= margin + footer_h:
                return
            draw_footer(page_number)
            canvas_obj.showPage()
            page_number += 1
            y = draw_table_header(draw_header(page_number))
            row_counter = 0

        for material_group, esp_group, group_rows in grouped_pieces:
            ensure_space(group_h + row_h)
            group_code = f"GRP|{of_code}|{material_group}|{esp_group}"
            group_y = y - group_h
            canvas_obj.setFillColor(colors.HexColor("#EEF6FF"))
            canvas_obj.setStrokeColor(colors.HexColor("#B9D7F2"))
            canvas_obj.roundRect(margin, group_y, inner_w, group_h - 1, 3, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.HexColor("#0F172A"))
            canvas_obj.setFont("Helvetica-Bold", 8.2)
            canvas_obj.drawString(margin + 3, group_y + 6.2 * mm, self._operator_pdf_text(f"{material_group} | {esp_group} mm"))
            canvas_obj.setFont("Helvetica", 6.4)
            canvas_obj.drawString(margin + 3, group_y + 2.4 * mm, self._operator_pdf_text(f"{len(group_rows)} peça(s) nesta espessura"))
            group_barcode_w = 104 * mm
            group_barcode_x = page_w - margin - group_barcode_w
            self._draw_code128_fit(canvas_obj, group_code, group_barcode_x, group_y + 2.2 * mm, group_barcode_w, 6.8 * mm, min_bar_width=0.34, max_bar_width=0.78)
            canvas_obj.setFont("Helvetica-Bold", 5.2)
            canvas_obj.drawCentredString(group_barcode_x + (group_barcode_w / 2), group_y + 0.7 * mm, self._operator_pdf_text(_pdf_clip_text(group_code, group_barcode_w, "Helvetica-Bold", 5.2)))
            y = group_y
            row_counter += 1
            for piece in group_rows:
                ensure_space(row_h)
                row_y = y - row_h
                canvas_obj.setFillColor(colors.HexColor("#FFFFFF") if row_counter % 2 == 0 else colors.HexColor("#F8FAFC"))
                canvas_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
                canvas_obj.rect(margin, row_y, inner_w, row_h, stroke=1, fill=1)
                values = [
                    str(piece.get("ref_interna", "-") or "-"),
                    str(piece.get("ref_externa", "-") or "-"),
                    str(piece.get("material", "-") or "-"),
                    str(piece.get("espessura", "-") or "-"),
                    str(piece.get("qtd_plan", "0") or "0"),
                    self._operations_pdf_abbrev(piece.get("operacoes", "-")),
                ]
                x = margin
                canvas_obj.setFillColor(colors.HexColor("#0F172A"))
                for col_index, value in enumerate(values):
                    width = columns[col_index][1]
                    font_name = "Helvetica-Bold" if col_index == 0 else "Helvetica"
                    font_size = 6.1 if col_index in {1, 5} else 6.5
                    canvas_obj.setFont(font_name, font_size)
                    if col_index == 5:
                        opp_txt = str(piece.get("opp", "") or "").strip()
                        barcode_w = min(width * 0.72, 64 * mm)
                        text_w = max(12 * mm, width - barcode_w - 5 * mm)
                        clipped = _pdf_clip_text(value, text_w, font_name, 5.7)
                        canvas_obj.setFont("Helvetica", 5.7)
                        canvas_obj.drawString(x + 2.2, row_y + 12.4 * mm, self._operator_pdf_text(clipped))
                        if opp_txt:
                            barcode_x = x + width - barcode_w - 2.2
                            self._draw_code128_fit(canvas_obj, opp_txt, barcode_x, row_y + 3.4 * mm, barcode_w, 6.4 * mm, min_bar_width=0.30, max_bar_width=0.72)
                            canvas_obj.setFont("Helvetica-Bold", 5.5)
                            canvas_obj.drawCentredString(barcode_x + (barcode_w / 2), row_y + 1.5 * mm, self._operator_pdf_text(_pdf_clip_text(opp_txt, barcode_w, "Helvetica-Bold", 5.5)))
                    elif col_index in {3, 4}:
                        clipped = _pdf_clip_text(value, width - 3.5, font_name, font_size)
                        canvas_obj.drawRightString(x + width - 2.2, row_y + 6.2 * mm, self._operator_pdf_text(clipped))
                    else:
                        clipped = _pdf_clip_text(value, width - 3.5, font_name, font_size)
                        canvas_obj.drawString(x + 2.2, row_y + 6.2 * mm, self._operator_pdf_text(clipped))
                    x += width
                y = row_y
                row_counter += 1
        draw_footer(page_number)
        for product_sheet in product_sheets:
            canvas_obj.showPage()
            page_number += 1
            draw_product_sheet(product_sheet, page_number)
        canvas_obj.save()
        extra_notice_pages: list[Path] = []
        if montagem_component_groups:
            component_pages: list[Path] = []

            def _start_component_page(page_idx: int, *, include_group_barcode: bool = False) -> tuple[Any, float, Path]:
                page_path = Path(tempfile.gettempdir()) / f"lugest_of_components_{of_code}_{page_idx}_{int(time.time() * 1000)}.pdf"
                page_canvas = pdf_canvas.Canvas(str(page_path), pagesize=A4)
                top_y = page_h - margin
                page_canvas.setFillColor(colors.white)
                page_canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
                page_canvas.roundRect(margin, top_y - 34 * mm, inner_w, 34 * mm, 6, stroke=1, fill=1)
                self._draw_operator_logo_plate(page_canvas, palette, logo_path, margin + 8, top_y - 20 * mm, 34 * mm, 15 * mm, radius=4, padding_x=3, padding_y=2)
                page_canvas.setFillColor(colors.HexColor("#020617"))
                page_canvas.setFont("Helvetica-Bold", 14)
                page_canvas.drawString(margin + 48 * mm, top_y - 11 * mm, self._operator_pdf_text("Componentes Montagem/Stock"))
                page_canvas.setFont("Helvetica", 8)
                page_canvas.drawString(margin + 48 * mm, top_y - 18 * mm, self._operator_pdf_text(f"OF: {of_code}"))
                page_canvas.drawString(margin + 48 * mm, top_y - 24 * mm, self._operator_pdf_text("Leitura individual por componente"))
                page_canvas.setStrokeColor(colors.HexColor("#E2E8F0"))
                page_canvas.line(margin, top_y - 39 * mm, page_w - margin, top_y - 39 * mm)
                section_y = top_y - 46 * mm
                if include_group_barcode and montagem_stock_code:
                    comp_h = 15 * mm
                    page_canvas.setFillColor(colors.HexColor("#F8FAFC"))
                    page_canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
                    page_canvas.roundRect(margin, section_y - comp_h, inner_w, comp_h, 4, stroke=1, fill=1)
                    page_canvas.setFillColor(colors.HexColor("#0F172A"))
                    page_canvas.setFont("Helvetica-Bold", 7.2)
                    page_canvas.drawString(margin + 3 * mm, section_y - 5 * mm, self._operator_pdf_text("Codigo do conjunto de componentes"))
                    page_canvas.setFont("Helvetica", 6.2)
                    page_canvas.drawString(margin + 3 * mm, section_y - 9.4 * mm, self._operator_pdf_text(f"{len(montagem_stock_items)} item(ns) de stock associados"))
                    self._draw_code128_fit(page_canvas, montagem_stock_code, page_w - margin - 74 * mm, section_y - 10.8 * mm, 70 * mm, 7.0 * mm, min_bar_width=0.28, max_bar_width=0.62)
                    page_canvas.setFont("Helvetica-Bold", 5.4)
                    page_canvas.drawCentredString(page_w - margin - 39 * mm, section_y - 12.5 * mm, self._operator_pdf_text(montagem_stock_code))
                    section_y -= comp_h + 4 * mm
                return page_canvas, section_y, page_path

            def _finish_component_page(page_canvas: Any, page_path: Path, page_idx: int) -> None:
                page_canvas.setStrokeColor(colors.HexColor("#E2E8F0"))
                page_canvas.line(margin, margin + 4.5 * mm, page_w - margin, margin + 4.5 * mm)
                page_canvas.setFillColor(colors.HexColor("#64748B"))
                page_canvas.setFont("Helvetica", 5.8)
                page_canvas.drawString(margin, margin + 1.4 * mm, self._operator_pdf_text(f"Impresso em: {printed_at}"))
                page_canvas.drawRightString(page_w - margin, margin + 1.4 * mm, self._operator_pdf_text(f"LUGEST | OF {of_code} | Componentes {page_idx}"))
                page_canvas.save()
                component_pages.append(page_path)

            page_index = 1
            component_page_has_rows = False
            first_component_page = True
            notice_canvas = None
            section_y = 0.0
            notice_path = None
            for esp_txt, rows in montagem_component_groups:
                needed_h = (10 + max(1, len(rows)) * 16) * mm
                if notice_canvas is None:
                    notice_canvas, section_y, notice_path = _start_component_page(
                        page_index,
                        include_group_barcode=first_component_page,
                    )
                if section_y - needed_h < margin + 16 * mm and component_page_has_rows:
                    _finish_component_page(notice_canvas, notice_path, page_index)
                    page_index += 1
                    first_component_page = False
                    notice_canvas, section_y, notice_path = _start_component_page(page_index)
                    component_page_has_rows = False
                notice_canvas.setFillColor(colors.HexColor("#EEF6FF"))
                notice_canvas.setStrokeColor(colors.HexColor("#B9D7F2"))
                notice_canvas.roundRect(margin, section_y - 8 * mm, inner_w, 8 * mm, 3, stroke=1, fill=1)
                notice_canvas.setFillColor(colors.HexColor("#0F172A"))
                notice_canvas.setFont("Helvetica-Bold", 8.3)
                notice_canvas.drawString(margin + 3 * mm, section_y - 5.4 * mm, self._operator_pdf_text(f"Espessura {esp_txt} mm"))
                section_y -= 11 * mm
                component_page_has_rows = True
                first_component_page = False
                for row in rows:
                    row_h_notice = 14 * mm
                    if section_y - row_h_notice < margin + 16 * mm:
                        _finish_component_page(notice_canvas, notice_path, page_index)
                        page_index += 1
                        notice_canvas, section_y, notice_path = _start_component_page(page_index)
                        component_page_has_rows = False
                    notice_canvas.setFillColor(colors.white)
                    notice_canvas.setStrokeColor(colors.HexColor("#D7DEE8"))
                    notice_canvas.roundRect(margin, section_y - row_h_notice, inner_w, row_h_notice, 3, stroke=1, fill=1)
                    notice_canvas.setFillColor(colors.HexColor("#0F172A"))
                    notice_canvas.setFont("Helvetica-Bold", 7.0)
                    notice_canvas.drawString(margin + 3 * mm, section_y - 4.6 * mm, self._operator_pdf_text(_pdf_clip_text(str(row.get("codigo", "-") or "-"), 52 * mm, "Helvetica-Bold", 7.0)))
                    notice_canvas.setFont("Helvetica", 6.0)
                    notice_canvas.drawString(margin + 3 * mm, section_y - 9.4 * mm, self._operator_pdf_text(_pdf_clip_text(str(row.get("descricao", "-") or "-"), 78 * mm, "Helvetica", 6.0)))
                    notice_canvas.setFont("Helvetica", 5.4)
                    notice_canvas.drawString(margin + 3 * mm, section_y - 12.8 * mm, self._operator_pdf_text(_pdf_clip_text(f"Qtd: {row.get('quantidade', '-')}", 42 * mm, "Helvetica", 5.4)))
                    barcode_x = page_w - margin - 72 * mm
                    notice_canvas.setFont("Helvetica-Bold", 5.2)
                    notice_canvas.drawString(barcode_x, section_y - 4.6 * mm, self._operator_pdf_text("Codigo individual"))
                    self._draw_code128_fit(notice_canvas, str(row.get("barcode", "") or "-"), barcode_x, section_y - 11.6 * mm, 68 * mm, 6.6 * mm, min_bar_width=0.24, max_bar_width=0.52)
                    notice_canvas.setFont("Helvetica-Bold", 4.8)
                    notice_canvas.drawCentredString(barcode_x + 34 * mm, section_y - 13.2 * mm, self._operator_pdf_text(_pdf_clip_text(str(row.get("barcode", "") or "-"), 68 * mm, "Helvetica-Bold", 4.8)))
                    section_y -= row_h_notice + 2 * mm
                    component_page_has_rows = True
                section_y -= 1 * mm
            if notice_canvas is not None and notice_path is not None:
                _finish_component_page(notice_canvas, notice_path, page_index)
            extra_notice_pages.extend(component_pages)
        if technical_groups or extra_notice_pages:
            try:
                from pypdf import PdfReader, PdfWriter

                writer = PdfWriter()

                def append_pdf(path_obj: Path) -> None:
                    reader = PdfReader(str(path_obj))
                    for page in reader.pages:
                        writer.add_page(page)

                def build_notice_page(title: str, lines: list[str]) -> Path:
                    notice_path = Path(tempfile.gettempdir()) / f"lugest_of_notice_{of_code}_{abs(hash(title + ''.join(lines)))}.pdf"
                    notice_canvas = pdf_canvas.Canvas(str(notice_path), pagesize=A4)
                    top_y = page_h - margin
                    notice_canvas.setFillColor(colors.white)
                    notice_canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
                    notice_canvas.roundRect(margin, top_y - 46 * mm, inner_w, 46 * mm, 6, stroke=1, fill=1)
                    self._draw_operator_logo_plate(notice_canvas, palette, logo_path, margin + 8, top_y - 24 * mm, 34 * mm, 15 * mm, radius=4, padding_x=3, padding_y=2)
                    notice_canvas.setFillColor(colors.HexColor("#020617"))
                    notice_canvas.setFont("Helvetica-Bold", 15)
                    notice_canvas.drawString(margin + 48 * mm, top_y - 12 * mm, self._operator_pdf_text(title))
                    notice_canvas.setFont("Helvetica", 8)
                    y_line = top_y - 25 * mm
                    for line in lines[:22]:
                        notice_canvas.drawString(margin + 10, y_line, self._operator_pdf_text(_pdf_clip_text(line, inner_w - 20, "Helvetica", 8)))
                        y_line -= 5 * mm
                    notice_canvas.save()
                    return notice_path

                def build_piece_cover(group: dict[str, Any], piece: dict[str, Any], docs: list[Path]) -> Path:
                    ref_internal = str(piece.get("ref_interna", "-") or "-").strip() or "-"
                    ref_external = str(piece.get("ref_externa", "-") or "-").strip() or "-"
                    material_txt = str(piece.get("material", group.get("material", "-")) or "-").strip() or "-"
                    esp_txt = str(piece.get("espessura", group.get("espessura", "-")) or "-").strip() or "-"
                    qty_txt = str(piece.get("qtd_plan", "0") or "0").strip() or "0"
                    ops_txt = str(piece.get("operacoes", "-") or "-").strip() or "-"
                    ops_short_txt = self._operations_pdf_abbrev(ops_txt)
                    weight_txt = self._fmt(piece.get("peso_unid", 0))
                    time_txt = self._fmt(piece.get("tempo_peca_min", 0))
                    dims_txt = str(piece.get("dimensao", "") or "").strip() or "-"
                    estado_txt = str(piece.get("estado", "") or "Em produção").strip() or "Em produção"
                    prioridade_txt = str(piece.get("prioridade", "") or "Normal").strip() or "Normal"
                    revisao_txt = str(piece.get("revisao", "") or "-").strip() or "-"
                    opp_txt = str(piece.get("opp", "") or "-").strip() or "-"
                    cover_path = Path(tempfile.gettempdir()) / f"lugest_of_piece_{of_code}_{abs(hash((ref_internal, ref_external, opp_txt, tuple(str(doc) for doc in docs))))}.pdf"
                    piece_canvas = pdf_canvas.Canvas(str(cover_path), pagesize=A4)
                    top_y = page_h - margin
                    black = colors.HexColor("#111111")
                    grey = colors.HexColor("#6B7280")
                    line = colors.HexColor("#9CA3AF")

                    def section_box(x: float, y_top: float, width: float, height: float, title: str) -> float:
                        piece_canvas.setFillColor(colors.white)
                        piece_canvas.setStrokeColor(line)
                        piece_canvas.roundRect(x, y_top - height, width, height, 3, stroke=1, fill=0)
                        piece_canvas.setFillColor(black)
                        piece_canvas.setFont("Helvetica-Bold", 10)
                        piece_canvas.drawString(x + 4 * mm, y_top - 7 * mm, self._operator_pdf_text(title.upper()))
                        piece_canvas.setStrokeColor(line)
                        piece_canvas.line(x + 4 * mm, y_top - 12 * mm, x + width - 4 * mm, y_top - 12 * mm)
                        return y_top - 19 * mm

                    def compact_section_box(x: float, y_top: float, width: float, height: float, title: str) -> float:
                        piece_canvas.setFillColor(colors.white)
                        piece_canvas.setStrokeColor(line)
                        piece_canvas.roundRect(x, y_top - height, width, height, 3, stroke=1, fill=0)
                        piece_canvas.setFillColor(black)
                        piece_canvas.setFont("Helvetica-Bold", 7.1)
                        piece_canvas.drawString(x + 3 * mm, y_top - 5 * mm, self._operator_pdf_text(title.upper()))
                        piece_canvas.setStrokeColor(line)
                        piece_canvas.line(x + 3 * mm, y_top - 8 * mm, x + width - 3 * mm, y_top - 8 * mm)
                        return y_top - 12 * mm

                    def detail_row(x: float, y_pos: float, width: float, label: str, value: str) -> float:
                        piece_canvas.setFillColor(black)
                        piece_canvas.setFont("Helvetica", 8.0)
                        piece_canvas.drawString(x, y_pos, self._operator_pdf_text(label))
                        piece_canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
                        piece_canvas.line(x + 22 * mm, y_pos - 1, x + width - 18 * mm, y_pos - 1)
                        value_width = max(18 * mm, width - 34 * mm)
                        value_size = _pdf_fit_font_size(str(value or "-"), "Helvetica", value_width, 8.3, 5.6)
                        piece_canvas.setFont("Helvetica", value_size)
                        piece_canvas.drawRightString(x + width, y_pos, self._operator_pdf_text(_pdf_clip_text(value, value_width, "Helvetica", value_size)))
                        return y_pos - 9 * mm

                    def bullet_lines(x: float, y_pos: float, width: float, lines: list[str], size: float = 8.5, max_count: int = 5) -> float:
                        piece_canvas.setFillColor(black)
                        piece_canvas.setFont("Helvetica", size)
                        for raw_line in lines[:max_count]:
                            text = _pdf_clip_text(str(raw_line or "-"), width - 7 * mm, "Helvetica", size)
                            piece_canvas.drawString(x, y_pos, self._operator_pdf_text("•"))
                            piece_canvas.drawString(x + 5 * mm, y_pos, self._operator_pdf_text(text))
                            y_pos -= 7 * mm
                        return y_pos

                    piece_canvas.setFillColor(colors.white)
                    piece_canvas.rect(0, 0, page_w, page_h, stroke=0, fill=1)

                    logo_x = margin
                    logo_y = top_y - 24 * mm
                    self._draw_operator_logo_plate(piece_canvas, palette, logo_path, logo_x, logo_y, 38 * mm, 17 * mm, radius=0, padding_x=0, padding_y=0)
                    piece_canvas.setStrokeColor(line)
                    piece_canvas.line(margin + 43 * mm, top_y - 5 * mm, margin + 43 * mm, top_y - 27 * mm)
                    piece_canvas.setFillColor(black)
                    title_max_w = 66 * mm
                    title_size = _pdf_fit_font_size("FOLHA DE ROSTO DA PEÇA", "Helvetica-Bold", title_max_w, 12.2, 9.4)
                    piece_canvas.setFont("Helvetica-Bold", title_size)
                    piece_canvas.drawString(margin + 50 * mm, top_y - 10 * mm, self._operator_pdf_text("FOLHA DE ROSTO DA PEÇA"))
                    piece_canvas.setFont("Helvetica-Bold", 9.2)
                    piece_canvas.drawString(margin + 50 * mm, top_y - 18 * mm, self._operator_pdf_text("ORDEM DE FABRICO"))

                    right_x = page_w - margin - 62 * mm
                    piece_canvas.line(right_x - 7 * mm, top_y - 4 * mm, right_x - 7 * mm, top_y - 32 * mm)
                    piece_canvas.setFont("Helvetica-Bold", 15)
                    piece_canvas.drawString(right_x, top_y - 9 * mm, self._operator_pdf_text(of_code))
                    piece_canvas.setFont("Helvetica-Bold", 9)
                    piece_canvas.drawString(right_x, top_y - 17 * mm, self._operator_pdf_text(f"{material_txt}  |  {esp_txt} mm"))
                    self._draw_code128_fit(piece_canvas, opp_txt, right_x, top_y - 29 * mm, 56 * mm, 8.2 * mm, min_bar_width=0.30, max_bar_width=0.70)
                    piece_canvas.setFont("Helvetica", 7)
                    piece_canvas.drawString(right_x, top_y - 34 * mm, self._operator_pdf_text(f"OPP: {opp_txt}"))

                    piece_canvas.setStrokeColor(line)
                    piece_canvas.line(margin, top_y - 43 * mm, page_w - margin, top_y - 43 * mm)

                    hero_top = top_y - 52 * mm
                    hero_h = 37 * mm
                    piece_canvas.setStrokeColor(line)
                    piece_canvas.roundRect(margin, hero_top - hero_h, inner_w, hero_h, 3, stroke=1, fill=0)
                    piece_canvas.setFillColor(black)
                    piece_canvas.setFont("Helvetica-Bold", 8)
                    piece_canvas.drawString(margin + 6 * mm, hero_top - 9 * mm, self._operator_pdf_text("CÓDIGO DA PEÇA"))
                    piece_canvas.setFont("Helvetica-Bold", 19)
                    piece_canvas.drawString(margin + 6 * mm, hero_top - 22 * mm, self._operator_pdf_text(_pdf_clip_text(ref_internal, 76 * mm, "Helvetica-Bold", 19)))
                    piece_canvas.setStrokeColor(line)
                    x1 = margin + 82 * mm
                    x2 = margin + 119 * mm
                    x3 = margin + 152 * mm
                    for sx in (x1, x2, x3):
                        piece_canvas.line(sx, hero_top - 5 * mm, sx, hero_top - hero_h + 5 * mm)
                    hero_items = [
                        (x1 + 9 * mm, "TEMPO PREVISTO", f"{time_txt} min", 8.2, 28 * mm, 2),
                        (x2 + 8 * mm, "PROCESSO", ops_short_txt, 8.0, 29 * mm, 3),
                        (x3 + 9 * mm, "PRIORIDADE", prioridade_txt, 8.2, 28 * mm, 2),
                    ]
                    for item_x, label, value, value_size, value_width, max_lines in hero_items:
                        piece_canvas.setFillColor(black)
                        piece_canvas.setFont("Helvetica", 7)
                        piece_canvas.drawCentredString(item_x + 14 * mm, hero_top - 13 * mm, self._operator_pdf_text(label))
                        piece_canvas.setFont("Helvetica-Bold", value_size)
                        for idx_line, wrapped in enumerate(_pdf_wrap_text(value, "Helvetica-Bold", value_size, value_width, max_lines=max_lines) or ["-"]):
                            piece_canvas.drawCentredString(item_x + 14 * mm, hero_top - (21 + (idx_line * 4.2)) * mm, self._operator_pdf_text(wrapped))

                    upper_top = hero_top - hero_h - 7 * mm
                    left_w = 90 * mm
                    right_w = inner_w - left_w - 5 * mm
                    data_body = section_box(margin, upper_top, left_w, 75 * mm, "Dados da peça")
                    y_info = data_body
                    for label, value in [
                        ("Ref. externa", ref_external),
                        ("Material", material_txt),
                        ("Espessura", f"{esp_txt} mm"),
                        ("Quantidade", f"{qty_txt} un"),
                        ("Peso unit.", f"{weight_txt} kg"),
                        ("Revisão", revisao_txt),
                        ("Prioridade", prioridade_txt),
                    ]:
                        y_info = detail_row(margin + 4 * mm, y_info, left_w - 8 * mm, label, value)

                    doc_body = section_box(margin + left_w + 5 * mm, upper_top, right_w, 38 * mm, "Documentação anexa")
                    doc_lines = [
                        f"Desenho técnico ({Path(docs[0]).name if docs else '-'})",
                        "Desenho de sequência / notas técnicas",
                        f"Total de PDF(s) associado(s): {len(docs)}",
                    ]
                    bullet_lines(margin + left_w + 9 * mm, doc_body, right_w - 12 * mm, doc_lines, size=8.4, max_count=3)

                    ops_top = upper_top - 81 * mm
                    ops_h = 47 * mm
                    ops_body = section_box(margin, ops_top, inner_w, ops_h, "Operações / picagem")
                    ops_lines = self._operations_pdf_legend(ops_txt)
                    op_names: list[str] = []
                    for legend_line in ops_lines:
                        legend_txt = str(legend_line or "").strip()
                        if not legend_txt:
                            continue
                        op_name = legend_txt.split(" - ", 1)[1].strip() if " - " in legend_txt else legend_txt
                        if op_name and op_name.lower() != "sem operações registadas" and op_name not in op_names:
                            op_names.append(op_name)
                    if op_names:
                        row_x = margin + 5 * mm
                        row_w = inner_w - 10 * mm
                        row_h = 7.4 * mm
                        row_top = ops_body + 2.0 * mm
                        max_ops = 5
                        for op_index, op_name in enumerate(op_names[:max_ops]):
                            cell_top = row_top - (op_index * row_h)
                            cell_y = cell_top - row_h + 1.0 * mm
                            op_abbrev = self._operation_pdf_abbrev(op_name)
                            op_code = f"OPR|{opp_txt}|{op_abbrev}"
                            piece_canvas.setStrokeColor(colors.HexColor("#D1D5DB"))
                            piece_canvas.setFillColor(colors.HexColor("#F8FAFC"))
                            piece_canvas.roundRect(row_x, cell_y, row_w, row_h - 0.7 * mm, 2, stroke=1, fill=1)
                            piece_canvas.setFillColor(black)
                            piece_canvas.setFont("Helvetica-Bold", 6.4)
                            piece_canvas.drawString(row_x + 2.6 * mm, cell_top - 3.0 * mm, self._operator_pdf_text(_pdf_clip_text(op_abbrev, 17 * mm, "Helvetica-Bold", 6.4)))
                            piece_canvas.setFont("Helvetica", 5.6)
                            piece_canvas.setFillColor(grey)
                            piece_canvas.drawString(row_x + 21 * mm, cell_top - 3.0 * mm, self._operator_pdf_text(_pdf_clip_text(op_name, 34 * mm, "Helvetica", 5.6)))
                            barcode_x = row_x + 58 * mm
                            barcode_w = row_w - 62 * mm
                            self._draw_code128_fit(piece_canvas, op_code, barcode_x, cell_y + 1.7 * mm, barcode_w, row_h - 3.1 * mm, min_bar_width=0.28, max_bar_width=0.62)
                            piece_canvas.setFont("Helvetica", 3.5)
                            piece_canvas.setFillColor(grey)
                            piece_canvas.drawCentredString(barcode_x + (barcode_w / 2), cell_y + 0.6 * mm, self._operator_pdf_text(_pdf_clip_text(op_code, barcode_w, "Helvetica", 3.5)))
                        if len(op_names) > max_ops:
                            piece_canvas.setFont("Helvetica", 5.5)
                            piece_canvas.setFillColor(grey)
                            piece_canvas.drawRightString(page_w - margin - 5 * mm, ops_top - ops_h + 4 * mm, self._operator_pdf_text(f"+{len(op_names) - max_ops} operação(ões)"))
                    else:
                        bullet_lines(margin + 5 * mm, ops_body, inner_w - 10 * mm, ops_lines, size=8.1, max_count=5)

                    lower_top = ops_top - ops_h - 6 * mm
                    third_gap = 5 * mm
                    third_w = (inner_w - (third_gap * 2)) / 3
                    lower_card_h = 28 * mm
                    prod_body = compact_section_box(margin, lower_top, third_w, lower_card_h, "Produção")
                    for line_txt in ["Posto:", "Operador:", "Data / Hora:"]:
                        piece_canvas.setFont("Helvetica", 6.4)
                        piece_canvas.setFillColor(black)
                        piece_canvas.drawString(margin + 4 * mm, prod_body, self._operator_pdf_text(line_txt))
                        piece_canvas.setStrokeColor(line)
                        piece_canvas.line(margin + 22 * mm, prod_body - 1, margin + third_w - 5 * mm, prod_body - 1)
                        prod_body -= 4.8 * mm

                    qual_body = compact_section_box(margin + third_w + third_gap, lower_top, third_w, lower_card_h, "Qualidade & controlo")
                    for line_txt in ["Conferir material e espessura", "Confirmar revisão do desenho", "Registar desvios ou não conformidades"]:
                        piece_canvas.circle(margin + third_w + third_gap + 7 * mm, qual_body + 0.8, 1.25 * mm, stroke=1, fill=0)
                        piece_canvas.setFont("Helvetica", 5.7)
                        piece_canvas.drawString(margin + third_w + third_gap + 11 * mm, qual_body, self._operator_pdf_text(_pdf_clip_text(line_txt, third_w - 15 * mm, "Helvetica", 5.7)))
                        qual_body -= 4.5 * mm

                    obs_body = compact_section_box(margin + (third_w + third_gap) * 2, lower_top, third_w, lower_card_h, "Observações")
                    obs_text = str(piece.get("descricao", "") or ref_external or "Sem observações.").strip()
                    for line_txt in _pdf_wrap_text(obs_text, "Helvetica", 5.9, third_w - 8 * mm, max_lines=2) or ["Sem observações."]:
                        piece_canvas.setFont("Helvetica", 5.9)
                        piece_canvas.drawString(margin + (third_w + third_gap) * 2 + 4 * mm, obs_body, self._operator_pdf_text(line_txt))
                        obs_body -= 4.5 * mm

                    trace_y = margin + 7 * mm
                    piece_canvas.setStrokeColor(line)
                    piece_canvas.line(margin, trace_y + 4 * mm, page_w - margin, trace_y + 4 * mm)
                    piece_canvas.setFont("Helvetica-Bold", 5.8)
                    piece_canvas.setFillColor(black)
                    piece_canvas.drawString(margin, trace_y, self._operator_pdf_text("RASTREABILIDADE"))
                    piece_canvas.setFont("Helvetica", 5.4)
                    piece_canvas.setFillColor(black)
                    trace_items = [
                        f"Dimensões: {dims_txt}",
                        f"Documento principal: {Path(docs[0]).name if docs else '-'}",
                        f"Material/Espessura: {material_txt} | {esp_txt} mm",
                    ]
                    trace_x = margin + 28 * mm
                    for idx_trace, value in enumerate(trace_items):
                        piece_canvas.drawString(trace_x + (idx_trace * 51 * mm), trace_y, self._operator_pdf_text(_pdf_clip_text(value, 47 * mm, "Helvetica", 5.4)))

                    piece_canvas.setStrokeColor(line)
                    piece_canvas.line(margin, margin + 4.5 * mm, page_w - margin, margin + 4.5 * mm)
                    piece_canvas.setFillColor(grey)
                    piece_canvas.setFont("Helvetica", 5.8)
                    piece_canvas.drawString(margin, margin + 1.4 * mm, self._operator_pdf_text(f"Impresso em: {printed_at}"))
                    piece_canvas.setFillColor(black)
                    piece_canvas.setFont("Helvetica", 6.2)
                    piece_canvas.drawRightString(page_w - margin, margin + 1.4 * mm, self._operator_pdf_text(f"LUGEST  |  OF: {of_code}  |  {ref_internal}"))
                    piece_canvas.save()
                    return cover_path

                append_pdf(cover_target)
                for group in technical_groups:
                    for doc_row in list(group.get("docs_rows", []) or []):
                        piece = dict(doc_row.get("piece", {}) or {})
                        docs_for_piece = [Path(doc_path) for doc_path in list(doc_row.get("docs", []) or [])]
                        append_pdf(build_piece_cover(group, piece, docs_for_piece))
                        for doc_path in docs_for_piece:
                            try:
                                append_pdf(Path(doc_path))
                            except Exception as exc:
                                append_pdf(
                                    build_notice_page(
                                        "PDF técnico não anexado",
                                        [
                                            f"Referência: {piece.get('ref_interna', '-') or '-'} | {piece.get('ref_externa', '-') or '-'}",
                                            f"Ficheiro: {Path(doc_path).name}",
                                            f"Caminho: {doc_path}",
                                            f"Erro: {exc}",
                                        ],
                                    )
                                )
                for page_path in extra_notice_pages:
                    append_pdf(page_path)
                with target.open("wb") as handle:
                    writer.write(handle)
            except Exception:
                try:
                    shutil.copy2(cover_target, target)
                except Exception:
                    pass
        return target

    def order_open_fabrication_pdf(self, numero: str, selected_groups: list[dict[str, Any]] | None = None) -> Path:
        path = self.order_fabrication_pdf(numero, selected_groups=selected_groups)
        try:
            os.startfile(str(path))
        except Exception:
            pass
        return path

    def order_print_fabrication_pdf(self, numero: str, selected_groups: list[dict[str, Any]] | None = None) -> Path:
        path = self.order_fabrication_pdf(numero, selected_groups=selected_groups)
        try:
            os.startfile(str(path))
        except Exception:
            pass
        return path
