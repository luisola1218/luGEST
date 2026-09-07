from __future__ import annotations

import re
import tempfile
from lugest_infra.pdf.dossier_reports import render_quality_dossier as _render_dossier_quality
from lugest_infra.pdf.text import (
    clip_text as _pdf_clip_text,
    fit_font_size as _pdf_fit_font_size,
    wrap_text as _pdf_wrap_text,
)
from pathlib import Path
from typing import Any


class QualityReportsBackendMixin:
    """Legacy adapter for quality reports; see BACKEND_GUIDE.md."""

    def _quality_pdf_path(self, name: str) -> Path:
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(name or "qualidade").strip()).strip("_") or "qualidade"
        return Path(tempfile.gettempdir()) / f"lugest_{safe}.pdf"

    def _quality_pdf_draw_lines(self, canvas_obj: Any, lines: list[str], x: float, y: float, width: float, *, size: float = 9.0) -> float:
        for raw in lines:
            for line in _pdf_wrap_text(raw, "Helvetica", size, width, max_lines=None) or [""]:
                canvas_obj.drawString(x, y, line)
                y -= size + 3
        return y

    def _quality_pdf_branding_assets(self) -> tuple[dict[str, Any], Path | None, list[str], str]:
        branding = self.branding_settings()
        palette = self._operator_label_palette()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        footer_lines = [str(line or "").strip() for line in list(branding.get("empresa_info_rodape", []) or []) if str(line or "").strip()]
        company_name = str(dict(branding.get("guia_emitente", {}) or {}).get("nome", "") or "").strip()
        if not company_name and footer_lines:
            company_name = footer_lines[0]
        if not company_name:
            company_name = "luGEST"
        return palette, logo_path, footer_lines[:3], company_name

    def _quality_pdf_draw_page_frame(
        self,
        canvas_obj: Any,
        page_w: float,
        page_h: float,
        *,
        title: str,
        subtitle: str = "",
        printed_at: str = "",
        page_label: str = "",
    ) -> dict[str, Any]:
        from reportlab.lib import colors
        from reportlab.lib.units import mm

        palette, logo_path, footer_lines, company_name = self._quality_pdf_branding_assets()
        outer_margin = 10 * mm
        header_h = 24 * mm
        footer_h = 18 * mm
        content_left = outer_margin + (8 * mm)
        content_right = page_w - outer_margin - (8 * mm)
        content_top = page_h - outer_margin - header_h - (6 * mm)
        content_bottom = outer_margin + footer_h + (6 * mm)

        canvas_obj.setFillColor(colors.white)
        canvas_obj.rect(0, 0, page_w, page_h, stroke=0, fill=1)
        canvas_obj.setFillColor(colors.white)
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_margin, outer_margin, page_w - (outer_margin * 2), page_h - (outer_margin * 2), 10, stroke=1, fill=1)

        header_y = page_h - outer_margin - header_h
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(0.8)
        canvas_obj.roundRect(outer_margin, header_y, page_w - (outer_margin * 2), header_h, 10, stroke=1, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            page_w - outer_margin - (34 * mm),
            header_y + (4.5 * mm),
            28 * mm,
            12 * mm,
            radius=5,
            padding_x=3,
            padding_y=2,
            line_width=0.7,
        )
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            page_w - outer_margin - (20 * mm),
            outer_margin + 3.5 * mm,
            14 * mm,
            8 * mm,
            radius=3,
            padding_x=2,
            padding_y=1.5,
            line_width=0.6,
        )
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont("Helvetica-Bold", 16)
        canvas_obj.drawString(content_left, header_y + (14.5 * mm), self._operator_pdf_text(title))
        if subtitle:
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont("Helvetica", 8.6)
            canvas_obj.drawString(content_left, header_y + (7.8 * mm), self._operator_pdf_text(_pdf_clip_text(subtitle, 120 * mm, "Helvetica", 8.6)))

        footer_y = outer_margin + 3.5 * mm
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(0.7)
        canvas_obj.line(content_left, outer_margin + footer_h, content_right, outer_margin + footer_h)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 7.2)
        footer_base = footer_y + 8.0
        footer_texts = footer_lines or [company_name]
        for index, line in enumerate(footer_texts[:2]):
            canvas_obj.drawString(content_left, footer_base + (index * 8.0), self._operator_pdf_text(_pdf_clip_text(line, 110 * mm, "Helvetica", 7.2)))
        if printed_at:
            canvas_obj.drawRightString(content_right, footer_base + 8.0, self._operator_pdf_text(printed_at))
        if page_label:
            canvas_obj.drawRightString(content_right, footer_base, self._operator_pdf_text(page_label))

        return {
            "palette": palette,
            "logo_path": logo_path,
            "content_left": content_left,
            "content_right": content_right,
            "content_top": content_top,
            "content_bottom": content_bottom,
        }

    def _quality_supplier_label_status_text(self, row: dict[str, Any], status_override: str = "") -> str:
        override = str(status_override or "").strip()
        if override:
            return override
        decision = str(row.get("decisao", "") or "").strip()
        decision_norm = decision.casefold()
        if "devolver" in decision_norm:
            return "DEVOLVER AO FORNECEDOR"
        if "aguardar" in decision_norm and "fornecedor" in decision_norm:
            return "AGUARDAR DECISAO DO FORNECEDOR"
        if "repor" in decision_norm or "substitu" in decision_norm:
            return "AGUARDAR REPOSICAO DO FORNECEDOR"
        if self._parse_float(row.get("qtd_rejeitada", 0), 0) > 0:
            return "REJEITADO"
        return decision.upper() or "EM ANALISE"

    def _quality_simple_pdf(self, target: Path, title: str, sections: list[tuple[str, list[str]]]) -> Path:
        def clean(value: Any) -> str:
            return str(value or "").replace("\r", " ").replace("\n", " ").strip()

        def esc(value: Any) -> str:
            text = clean(value)
            text = text.encode("latin-1", errors="replace").decode("latin-1")
            return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

        all_lines: list[str] = [clean(title), ""]
        for section_title, lines in sections:
            all_lines.append(clean(section_title))
            all_lines.extend(clean(line) for line in list(lines or []))
            all_lines.append("")
        wrapped: list[str] = []
        for line in all_lines:
            if not line:
                wrapped.append("")
                continue
            text = line
            while len(text) > 96:
                wrapped.append(text[:96])
                text = text[96:]
            wrapped.append(text)
        page_lines: list[list[str]] = []
        current: list[str] = []
        for line in wrapped:
            current.append(line)
            if len(current) >= 48:
                page_lines.append(current)
                current = []
        if current or not page_lines:
            page_lines.append(current)

        objects: list[bytes] = []
        pages_obj_num = 2
        font_obj_num = 3
        page_obj_nums: list[int] = []
        content_obj_nums: list[int] = []
        next_obj = 4
        for _page in page_lines:
            page_obj_nums.append(next_obj)
            content_obj_nums.append(next_obj + 1)
            next_obj += 2
        objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
        kids = " ".join(f"{num} 0 R" for num in page_obj_nums)
        objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {len(page_obj_nums)} >>".encode("ascii"))
        objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        for page_num, content_num, lines in zip(page_obj_nums, content_obj_nums, page_lines):
            objects.append(f"<< /Type /Page /Parent {pages_obj_num} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 {font_obj_num} 0 R >> >> /Contents {content_num} 0 R >>".encode("ascii"))
            stream_lines = ["BT", "/F1 10 Tf", "50 800 Td", "14 TL"]
            for line in lines:
                stream_lines.append(f"({esc(line)}) Tj")
                stream_lines.append("T*")
            stream_lines.append("ET")
            stream = "\n".join(stream_lines).encode("latin-1", errors="replace")
            objects.append(b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream")

        payload = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for index, obj in enumerate(objects, start=1):
            offsets.append(len(payload))
            payload.extend(f"{index} 0 obj\n".encode("ascii"))
            payload.extend(obj)
            payload.extend(b"\nendobj\n")
        xref_at = len(payload)
        payload.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
        for offset in offsets[1:]:
            payload.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
        payload.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_at}\n%%EOF\n".encode("ascii"))
        target.write_bytes(bytes(payload))
        return target

    def quality_nc_pdf(self, nc_id: str) -> Path:
        nc_id_txt = str(nc_id or "").strip()
        row = next((item for item in self.quality_nc_rows("", "Todos") if str(item.get("id", "") or "").strip() == nc_id_txt), None)
        if not row:
            raise ValueError("Nao conformidade nao encontrada.")
        target = self._quality_pdf_path(f"NC_{nc_id_txt}")
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except Exception:
            path = self._quality_simple_pdf(
                target,
                f"Nao conformidade {nc_id_txt}",
                [
                    ("Identificacao", [f"{key}: {row.get(key, '')}" for key in ("estado", "gravidade", "tipo", "origem", "referencia", "entidade_label", "fornecedor_nome", "lote_fornecedor", "ne_numero", "responsavel", "prazo")]),
                    ("Descricao", [row.get("descricao", "") or "-"]),
                    ("Causa", [row.get("causa", "") or "-"]),
                    ("Acao corretiva", [row.get("acao", "") or "-"]),
                    ("Eficacia", [row.get("eficacia", "") or "-"]),
                ],
            )
            return path
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=A4)
        page_w, page_h = A4
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        page_number = 1

        def begin_page() -> dict[str, Any]:
            return self._quality_pdf_draw_page_frame(
                canvas_obj,
                page_w,
                page_h,
                title="Nao Conformidade",
                subtitle=f"{nc_id_txt} | {str(row.get('estado', '') or '-').strip()} | {str(row.get('gravidade', '') or '-').strip()}",
                printed_at=printed_at,
                page_label=f"Pagina {page_number}",
            )

        frame = begin_page()
        palette = frame["palette"]
        content_left = float(frame["content_left"])
        content_right = float(frame["content_right"])
        content_top = float(frame["content_top"])
        bottom_limit = float(frame["content_bottom"])
        y = content_top

        def clean(value: Any) -> str:
            return str(value or "-").replace("\r", " ").replace("\n", " ").strip() or "-"

        def next_page() -> None:
            nonlocal page_number, frame, palette, content_left, content_right, content_top, bottom_limit, y
            canvas_obj.showPage()
            page_number += 1
            frame = begin_page()
            palette = frame["palette"]
            content_left = float(frame["content_left"])
            content_right = float(frame["content_right"])
            content_top = float(frame["content_top"])
            bottom_limit = float(frame["content_bottom"])
            y = content_top

        def ensure_space(required_height: float) -> None:
            if y - required_height < bottom_limit:
                next_page()

        def card(x: float, top: float, width: float, height: float, title: str) -> float:
            canvas_obj.setFillColor(palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.setLineWidth(0.8)
            canvas_obj.roundRect(x, top - height, width, height, 7, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["surface_alt"])
            canvas_obj.roundRect(x, top - 17, width, 17, 7, stroke=0, fill=1)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont("Helvetica-Bold", 8.8)
            canvas_obj.drawString(x + 8, top - 11.5, self._operator_pdf_text(title))
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.line(x + 8, top - 22, x + width - 8, top - 22)
            return top - 30

        def draw_kpi(x: float, top: float, width: float, label: str, value: Any, accent: Any | None = None) -> None:
            canvas_obj.setFillColor(palette["surface_alt"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(x, top - 34, width, 34, 7, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont("Helvetica", 6.4)
            canvas_obj.drawString(x + 7, top - 10, self._operator_pdf_text(label))
            value_txt = clean(value)
            value_font = _pdf_fit_font_size(value_txt, "Helvetica-Bold", width - 14, 11.2, 7.2)
            canvas_obj.setFillColor(accent or palette["ink"])
            canvas_obj.setFont("Helvetica-Bold", value_font)
            canvas_obj.drawString(x + 7, top - 23.5, self._operator_pdf_text(_pdf_clip_text(value_txt, width - 14, "Helvetica-Bold", value_font)))

        def draw_pair(label: str, value: Any, x: float, y_pos: float, width: float) -> None:
            label_w = min(34 * mm, width * 0.38)
            value_w = max(20.0, width - label_w - 8)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont("Helvetica", 7.2)
            canvas_obj.drawString(x, y_pos, self._operator_pdf_text(label))
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.line(x + label_w, y_pos - 1.5, x + label_w + value_w, y_pos - 1.5)
            value_txt = clean(value)
            value_font = _pdf_fit_font_size(value_txt, "Helvetica-Bold", value_w, 8.0, 6.0)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont("Helvetica-Bold", value_font)
            canvas_obj.drawRightString(x + label_w + value_w, y_pos, self._operator_pdf_text(_pdf_clip_text(value_txt, value_w, "Helvetica-Bold", value_font)))

        status = clean(row.get("estado", ""))
        severity = clean(row.get("gravidade", ""))
        decision = clean(row.get("decisao", ""))
        accent = palette["primary_dark"]
        if any(token in f"{status} {severity} {decision}".casefold() for token in ("rejeit", "crit", "devolver")):
            accent = palette["danger"]
        elif any(token in f"{status} {severity} {decision}".casefold() for token in ("aguard", "pend", "anal")):
            accent = palette["warning"]

        kpi_gap = 6
        kpi_w = (content_right - content_left - (kpi_gap * 3)) / 4.0
        for index, (label, key) in enumerate((("NC", "id"), ("Estado", "estado"), ("Gravidade", "gravidade"), ("Tipo", "tipo"))):
            draw_kpi(content_left + (index * (kpi_w + kpi_gap)), y, kpi_w, label, row.get(key, ""), accent if index in (1, 2) else None)
        y -= 44

        main_gap = 8 * mm
        left_w = (content_right - content_left - main_gap) * 0.56
        right_w = content_right - content_left - main_gap - left_w
        row_top = y
        card_y = card(content_left, row_top, left_w, 96, "Identificacao e origem")
        for label, key in (
            ("Referencia", "referencia"),
            ("Entidade", "entidade_label"),
            ("Fornecedor", "fornecedor_nome"),
            ("NE / Guia", "ne_numero"),
            ("Lote fornecedor", "lote_fornecedor"),
        ):
            draw_pair(label, row.get(key, ""), content_left + 8, card_y, left_w - 16)
            card_y -= 12

        right_x = content_left + left_w + main_gap
        card_y = card(right_x, row_top, right_w, 96, "Quantidades")
        qty_pairs = [
            ("Recebida", row.get("qtd_recebida", "")),
            ("Aprovada", row.get("qtd_aprovada", "")),
            ("Rejeitada", row.get("qtd_rejeitada", "")),
            ("Pendente", row.get("qtd_pendente", "")),
        ]
        for label, value in qty_pairs:
            draw_pair(label, value, right_x + 8, card_y, right_w - 16)
            card_y -= 13
        y = row_top - 108

        meta_top = y
        half_w = (content_right - content_left - main_gap) / 2.0
        card_y = card(content_left, meta_top, half_w, 60, "Responsabilidade")
        draw_pair("Responsavel", row.get("responsavel", ""), content_left + 8, card_y, half_w - 16)
        draw_pair("Prazo", row.get("prazo", ""), content_left + 8, card_y - 13, half_w - 16)
        draw_pair("Criada em", row.get("created_at", ""), content_left + 8, card_y - 26, half_w - 16)

        card_y = card(content_left + half_w + main_gap, meta_top, half_w, 60, "Decisao")
        decision_lines = _pdf_wrap_text(decision, "Helvetica-Bold", 8.0, half_w - 18, max_lines=3) or ["-"]
        canvas_obj.setFillColor(accent)
        canvas_obj.setFont("Helvetica-Bold", 8.0)
        line_y = card_y
        for line in decision_lines[:3]:
            canvas_obj.drawString(content_left + half_w + main_gap + 8, line_y, self._operator_pdf_text(line))
            line_y -= 10
        y = meta_top - 72

        def draw_text_section(title: str, value: Any, min_h: float = 54) -> None:
            nonlocal y
            text = clean(value)
            width = content_right - content_left
            lines = _pdf_wrap_text(text, "Helvetica", 8.6, width - 18, max_lines=None) or ["-"]
            height = max(min_h, 34 + (len(lines) * 10))
            ensure_space(height + 8)
            text_y = card(content_left, y, width, height, title)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont("Helvetica", 8.6)
            for line in lines:
                if text_y < bottom_limit + 12:
                    next_page()
                    text_y = card(content_left, y, width, height, title + " (cont.)")
                    canvas_obj.setFillColor(palette["ink"])
                    canvas_obj.setFont("Helvetica", 8.6)
                canvas_obj.drawString(content_left + 8, text_y, self._operator_pdf_text(line))
                text_y -= 10
            y -= height + 8

        draw_text_section("Descricao da nao conformidade", row.get("descricao", ""), 62)
        draw_text_section("Causa provavel", row.get("causa", ""), 52)
        draw_text_section("Acao corretiva / contencao", row.get("acao", ""), 58)
        draw_text_section("Eficacia / fecho", row.get("eficacia", ""), 48)

        canvas_obj.save()
        return target

    def quality_supplier_label_pdf(self, nc_id: str, output_path: str | Path | None = None, status_override: str = "") -> Path:
        nc_id_txt = str(nc_id or "").strip()
        row = next((item for item in self.quality_nc_rows("", "Todos") if str(item.get("id", "") or "").strip() == nc_id_txt), None)
        if not row:
            raise ValueError("Nao conformidade nao encontrada.")
        try:
            from reportlab.lib.pagesizes import A5
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas as pdf_canvas
        except Exception:
            target = Path(output_path) if output_path else self._quality_pdf_path(f"etiqueta_fornecedor_{nc_id_txt}")
            return self._quality_simple_pdf(
                target,
                f"Etiqueta fornecedor {nc_id_txt}",
                [
                    (
                        "Resumo",
                        [
                            f"Estado etiqueta: {self._quality_supplier_label_status_text(row, status_override)}",
                            f"Fornecedor: {row.get('fornecedor_nome', '') or '-'}",
                            f"Referencia: {row.get('referencia', '') or '-'}",
                            f"Lote: {row.get('lote_fornecedor', '') or '-'}",
                            f"Qtd rejeitada: {row.get('qtd_rejeitada', '') or '-'}",
                            f"Decisao: {row.get('decisao', '') or '-'}",
                        ],
                    )
                ],
            )

        target = (
            Path(output_path)
            if output_path
            else self._storage_output_path("quality/labels", f"Etiqueta_Qualidade_{nc_id_txt}.pdf")
        )
        target.parent.mkdir(parents=True, exist_ok=True)
        page_w, page_h = A5
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(page_w, page_h))
        palette, logo_path, _footer_lines, _company_name = self._quality_pdf_branding_assets()
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        status_text = self._quality_supplier_label_status_text(row, status_override)
        status_norm = status_text.casefold()
        status_color = palette["warning"]
        if "rejeitado" in status_norm or "devolver" in status_norm:
            status_color = palette["danger"]
        elif "repos" in status_norm or "substit" in status_norm:
            status_color = palette["primary_dark"]

        trim_margin = 6 * mm
        safe_margin = 11 * mm
        outer_x = safe_margin
        outer_y = safe_margin
        outer_w = page_w - (safe_margin * 2)
        outer_h = page_h - (safe_margin * 2)
        header_h = 31 * mm
        body_x = outer_x + 10
        body_w = outer_w - 20
        status_box_w = 46 * mm
        status_box_h = 22 * mm

        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(0.8)
        canvas_obj.roundRect(trim_margin, trim_margin, page_w - (trim_margin * 2), page_h - (trim_margin * 2), 7, stroke=1, fill=0)
        canvas_obj.setDash(3, 2)
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(outer_x - 4, outer_y - 4, outer_w + 8, outer_h + 8, 9, stroke=1, fill=0)
        canvas_obj.setDash()

        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.setLineWidth(1)
        canvas_obj.roundRect(outer_x, outer_y, outer_w, outer_h, 11, stroke=1, fill=1)

        header_y = outer_y + outer_h - header_h
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(outer_x + 8, header_y + 6, outer_w - 16, header_h - 12, 9, stroke=1, fill=1)
        self._draw_operator_logo_plate(
            canvas_obj,
            palette,
            logo_path,
            outer_x + 9,
            outer_y + outer_h - (18 * mm),
            24 * mm,
            10 * mm,
            radius=5,
            padding_x=3,
            padding_y=2,
            line_width=0.7,
        )
        title_x = outer_x + 36 * mm
        title_w = outer_w - (title_x - outer_x) - status_box_w - 16
        canvas_obj.setFillColor(palette["ink"])
        title_font = _pdf_fit_font_size("Etiqueta fornecedor", "Helvetica-Bold", title_w, 14.2, 11.4)
        canvas_obj.setFont("Helvetica-Bold", title_font)
        canvas_obj.drawString(title_x, outer_y + outer_h - (10.1 * mm), self._operator_pdf_text("Etiqueta fornecedor"))
        subtitle = f"NC {nc_id_txt} | Segregacao / devolucao | formato A5"
        subtitle_font = _pdf_fit_font_size(subtitle, "Helvetica", title_w, 7.6, 6.1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", subtitle_font)
        canvas_obj.drawString(title_x, outer_y + outer_h - (14.4 * mm), self._operator_pdf_text(_pdf_clip_text(subtitle, title_w, "Helvetica", subtitle_font)))

        status_box_x = outer_x + outer_w - status_box_w - 10
        status_box_y = outer_y + outer_h - header_h + 1
        canvas_obj.setFillColor(palette["surface_alt"])
        canvas_obj.setStrokeColor(status_color)
        canvas_obj.setLineWidth(1.1)
        canvas_obj.roundRect(status_box_x, status_box_y, status_box_w, status_box_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 6.1)
        canvas_obj.drawString(status_box_x + 7, status_box_y + status_box_h - 8, self._operator_pdf_text("Estado"))
        canvas_obj.setFillColor(status_color)
        status_lines = _pdf_wrap_text(status_text, "Helvetica-Bold", 9.0, status_box_w - 14, max_lines=2) or [status_text]
        status_y = status_box_y + status_box_h - 16
        status_font = _pdf_fit_font_size(max(status_lines, key=len), "Helvetica-Bold", status_box_w - 14, 8.8, 6.8)
        canvas_obj.setFont("Helvetica-Bold", status_font)
        for line in status_lines[:2]:
            canvas_obj.drawCentredString(status_box_x + (status_box_w / 2.0), status_y, self._operator_pdf_text(line))
            status_y -= status_font + 1.0

        body_top = outer_y + outer_h - header_h - 10
        left_w = body_w - status_box_w - 8
        row_y = body_top - 18
        small_gap = 6
        small_w = (left_w - small_gap) / 2.0
        info_cards = [
            ("Fornecedor", str(row.get("fornecedor_nome", "") or row.get("entidade_label", "") or "-").strip() or "-"),
            ("Lote", str(row.get("lote_fornecedor", "") or "-").strip() or "-"),
            ("Referencia", str(row.get("referencia", "") or "-").strip() or "-"),
            ("Qtd rejeitada", self._fmt(row.get("qtd_rejeitada", 0))),
        ]
        for index, (label, value) in enumerate(info_cards):
            col = index % 2
            line = index // 2
            box_x = body_x + (col * (small_w + small_gap))
            box_y = row_y - (line * 27)
            canvas_obj.setFillColor(palette["surface_alt"] if line == 0 else palette["surface"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(box_x, box_y, small_w, 22, 7, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont("Helvetica", 6.0)
            canvas_obj.drawString(box_x + 7, box_y + 12.6, self._operator_pdf_text(label))
            value_font = _pdf_fit_font_size(value, "Helvetica-Bold", small_w - 14, 8.0, 6.2)
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont("Helvetica-Bold", value_font)
            canvas_obj.drawString(box_x + 7, box_y + 4.2, self._operator_pdf_text(_pdf_clip_text(value, small_w - 14, "Helvetica-Bold", value_font)))

        decision_y = row_y - 60
        decision_text = str(row.get("decisao", "") or "-").strip() or "-"
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, decision_y, body_w, 22, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 6.1)
        canvas_obj.drawString(body_x + 8, decision_y + 13.0, self._operator_pdf_text("Instrucao / decisao"))
        decision_font = _pdf_fit_font_size(decision_text, "Helvetica-Bold", body_w - 16, 8.4, 6.5)
        canvas_obj.setFillColor(palette["ink"])
        canvas_obj.setFont("Helvetica-Bold", decision_font)
        canvas_obj.drawString(body_x + 8, decision_y + 4.2, self._operator_pdf_text(_pdf_clip_text(decision_text, body_w - 16, "Helvetica-Bold", decision_font)))

        barcode_y = outer_y + 10
        barcode_h = max(34.0, decision_y - barcode_y - 8)
        barcode_value = nc_id_txt
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(body_x, barcode_y, body_w, barcode_h, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 6.0)
        canvas_obj.drawString(body_x + 8, barcode_y + barcode_h - 11.0, self._operator_pdf_text("Codigo para rastreabilidade"))
        barcode_area_x = body_x + 10
        barcode_area_w = body_w - 20
        barcode_bar_h = max(17.0, min(24.0, barcode_h - 21.0))
        self._draw_code128_fit(canvas_obj, barcode_value, barcode_area_x, barcode_y + 11.0, barcode_area_w, barcode_bar_h, min_bar_width=0.42, max_bar_width=1.0)
        canvas_obj.setFillColor(palette["ink"])
        barcode_font = _pdf_fit_font_size(barcode_value, "Helvetica-Bold", barcode_area_w, 8.0, 6.0)
        canvas_obj.setFont("Helvetica-Bold", barcode_font)
        canvas_obj.drawCentredString(body_x + (body_w / 2.0), barcode_y + 4.0, self._operator_pdf_text(_pdf_clip_text(barcode_value, barcode_area_w, "Helvetica-Bold", barcode_font)))
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont("Helvetica", 5.2)
        canvas_obj.drawRightString(outer_x + outer_w - 8, outer_y + 4.8, self._operator_pdf_text(printed_at[:16]))
        canvas_obj.save()
        self._append_audit_event(self.ensure_data(), action="Etiqueta fornecedor gerada", entity_type="Nao conformidade", entity_id=nc_id_txt, summary=str(target))
        self._save(force=True, audit=False)
        return target

    def quality_dossier_pdf(self) -> Path:
        target = self._quality_pdf_path("dossier_qualidade")
        path = _render_dossier_quality(self, target)
        self._append_audit_event(self.ensure_data(), action="Dossier qualidade gerado", entity_type="Qualidade", entity_id="ISO9001", summary=str(path))
        self._save(force=True, audit=False)
        return path

    def quality_iso_checklist(self) -> list[dict[str, str]]:
        summary = self.quality_summary()
        docs = self.quality_document_rows()
        audit_count = int(summary.get("audit_events", 0) or 0)
        open_nc = int(summary.get("open_nc", 0) or 0)
        overdue_nc = int(summary.get("overdue_nc", 0) or 0)
        has_docs = bool(docs)
        blocked_materials = int(summary.get("blocked_materials", 0) or 0)
        supplier_nc = int(summary.get("supplier_nc", 0) or 0)
        return [
            {"area": "Rastreabilidade", "estado": "OK" if audit_count > 0 else "Pendente", "evidencia": f"{audit_count} eventos de auditoria registados."},
            {"area": "Nao conformidades", "estado": "Atencao" if overdue_nc else "OK", "evidencia": f"{open_nc} abertas; {overdue_nc} fora de prazo."},
            {"area": "Rececao e stock", "estado": "Atencao" if blocked_materials else "OK", "evidencia": f"{blocked_materials} lotes bloqueados/em inspecao com rastreabilidade."},
            {"area": "Reclamacoes a fornecedores", "estado": "OK" if supplier_nc or not blocked_materials else "Pendente", "evidencia": f"{supplier_nc} NC/reclamacoes de fornecedor registadas."},
            {"area": "Informacao documentada", "estado": "OK" if has_docs else "Pendente", "evidencia": f"{len(docs)} documentos/evidencias ligados ao sistema."},
            {"area": "Integridade das ligacoes", "estado": "OK" if int(summary.get("quality_issues", 0) or 0) == 0 else "Atencao", "evidencia": f"{int(summary.get('quality_issues', 0) or 0)} problemas em NC/documentos ligados."},
            {"area": "Alteracoes climaticas ISO 9001:2015/Amd 1:2024", "estado": "Pendente", "evidencia": "Registar no contexto da organizacao se o tema e relevante e que requisitos de partes interessadas existem."},
        ]
