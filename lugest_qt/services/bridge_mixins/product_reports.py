from __future__ import annotations

import os
import re
import tempfile
from lugest_infra.pdf.font_policy import pdf_font_size_limit
from lugest_infra.pdf.light_inventory import render_product_stock_pdf as _render_light_product_stock_pdf
from lugest_infra.pdf.light_labels import draw_product_label as _draw_light_product_label
from lugest_infra.pdf.text import clip_text as _pdf_clip_text, wrap_text as _pdf_wrap_text
from pathlib import Path
from typing import Any


class ProductReportsBackendMixin:
    """Legacy adapter for product reports; see BACKEND_GUIDE.md."""

    def product_stock_filters(self) -> dict[str, list[str]]:
        rows = [row for row in list(self.ensure_data().get("produtos", []) or []) if isinstance(row, dict)]
        return {
            "categories": sorted(
                {str(row.get("categoria", "") or "Sem categoria").strip() for row in rows},
                key=str.casefold,
            ),
            "types": sorted(
                {str(row.get("tipo", "") or "Sem tipo").strip() for row in rows},
                key=str.casefold,
            ),
        }

    def product_render_stock_pdf(
        self,
        path: str | Path,
        categories: list[str] | None = None,
        types: list[str] | None = None,
        in_stock_only: bool = False,
    ) -> Path:
        return _render_light_product_stock_pdf(
            Path(path),
            self.ensure_data(),
            self.branding_settings(),
            self.desktop_main.produto_preco_unitario,
            categories=categories,
            types=types,
            in_stock_only=bool(in_stock_only),
        )

    def product_open_stock_pdf(
        self,
        categories: list[str] | None = None,
        types: list[str] | None = None,
        in_stock_only: bool = False,
    ) -> Path:
        target = Path(tempfile.gettempdir()) / "lugest_qt_produtos_stock.pdf"
        self.product_render_stock_pdf(
            target,
            categories=categories,
            types=types,
            in_stock_only=in_stock_only,
        )
        os.startfile(str(target))
        return target

    def product_sheet_pdf(self, codigo: str, output_path: str | Path | None = None) -> Path:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfgen import canvas as pdf_canvas

        product = self.product_detail(codigo)
        code = str(product.get("codigo", "") or "").strip() or "produto"
        scan_code = str(product.get("scan_code", "") or self.inventory_scan_code("PRD", code)).strip()
        safe_code = re.sub(r"[^A-Za-z0-9_.-]+", "_", code).strip("_") or "produto"
        target = (
            Path(output_path)
            if output_path
            else self._storage_output_path("products/sheets", f"Ficha_Produto_{safe_code}.pdf")
        )
        target.parent.mkdir(parents=True, exist_ok=True)

        page_w, page_h = A4
        margin = 34
        palette = self._operator_label_palette()
        regular_font = "Helvetica"
        bold_font = "Helvetica-Bold"
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None

        qty = self._parse_float(product.get("qty", 0), 0)
        available = self._parse_float(product.get("available_qty", product.get("qty", 0)), 0)
        unit = str(product.get("unid", "") or "UN").strip() or "UN"
        price_unit = self._parse_float(product.get("preco_unid", 0), 0)
        stock_value = self._parse_float(product.get("valor_stock", qty * price_unit), 0)
        alert = self._parse_float(product.get("alerta", 0), 0)
        state = "Sem stock" if available <= 0 else ("Stock baixo" if alert > 0 and available <= alert else "Disponivel")
        desc = str(product.get("descricao", "") or "-").strip() or "-"
        category = str(product.get("categoria", "") or "-").strip() or "-"
        subcat = str(product.get("subcat", "") or "-").strip() or "-"
        kind = str(product.get("tipo", "") or "-").strip() or "-"
        dim_text = self._product_dimensoes(product)
        movements = list(product.get("movimentos", []) or [])[:8]

        def txt(value: Any) -> str:
            return self._operator_pdf_text(value)

        def money(value: Any) -> str:
            return f"{self._fmt(value)} EUR"

        def fit(text: str, max_w: float, preferred: float, minimum: float = 7.0, font: str = bold_font) -> float:
            size = float(preferred)
            while size > minimum and pdfmetrics.stringWidth(str(text or ""), font, size) > max_w:
                size -= 0.25
            return max(minimum, size)

        def chip(canvas_obj: Any, x: float, y: float, w: float, h: float, label: str, value: str, accent: bool = False) -> None:
            canvas_obj.setFillColor(palette["primary_soft"] if accent else palette["surface"])
            canvas_obj.setStrokeColor(palette["line_strong"] if accent else palette["line"])
            canvas_obj.roundRect(x, y, w, h, 8, stroke=1, fill=1)
            compact = h <= 34
            label_size = 6.5 if compact else 7.2
            label_y = y + h - (10.5 if compact else 13)
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(regular_font, label_size)
            canvas_obj.drawString(x + 9, label_y, txt(_pdf_clip_text(label, w - 18, regular_font, label_size)))
            value_font = fit(value, w - 18, 10.5 if compact else 13.5, 7.0 if compact else 8.0)
            value_y = y + (5.5 if compact else 9.5)
            canvas_obj.setFillColor(palette["primary_dark"] if accent else palette["ink"])
            canvas_obj.setFont(bold_font, value_font)
            canvas_obj.drawString(x + 9, value_y, txt(_pdf_clip_text(value, w - 18, bold_font, value_font)))

        def section_title(canvas_obj: Any, title: str, y: float) -> None:
            canvas_obj.setFillColor(palette["primary_dark"])
            canvas_obj.setFont(bold_font, 11.2)
            canvas_obj.drawString(margin, y, txt(title))
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.line(margin, y - 5, page_w - margin, y - 5)

        c = pdf_canvas.Canvas(str(target), pagesize=A4)
        c.setTitle(txt(f"Ficha de Produto {code}"))
        c.setFillColor(palette["surface_alt"])
        c.rect(0, 0, page_w, page_h, stroke=0, fill=1)

        header_h = 106
        c.setFillColor(palette["surface"])
        c.setStrokeColor(palette["line"])
        c.roundRect(margin, page_h - margin - header_h, page_w - (2 * margin), header_h, 14, stroke=1, fill=1)
        self._draw_operator_logo_plate(c, palette, logo_path, margin + 14, page_h - margin - 54, 104, 40, radius=8, padding_x=5, padding_y=4)
        title_x = margin + 134
        title_w = page_w - margin - title_x - 118
        c.setFillColor(palette["muted"])
        c.setFont(regular_font, 8)
        c.drawString(title_x, page_h - margin - 28, txt("Ficha tecnica e comercial"))
        code_font = fit(code, title_w, 22, 13)
        c.setFillColor(palette["primary_dark"])
        c.setFont(bold_font, code_font)
        c.drawString(title_x, page_h - margin - 52, txt(_pdf_clip_text(code, title_w, bold_font, code_font)))
        desc_lines = _pdf_wrap_text(desc, regular_font, 9.2, title_w, max_lines=2) or ["-"]
        c.setFillColor(palette["ink"])
        c.setFont(regular_font, 9.2)
        for line_index, line in enumerate(desc_lines):
            c.drawString(title_x, page_h - margin - 72 - (line_index * 11), txt(line))

        status_x = page_w - margin - 100
        c.setFillColor(palette["primary_soft"])
        c.setStrokeColor(palette["line_strong"])
        c.roundRect(status_x, page_h - margin - 58, 86, 34, 10, stroke=1, fill=1)
        c.setFillColor(palette["muted"])
        c.setFont(regular_font, 7)
        c.drawString(status_x + 9, page_h - margin - 36, txt("Estado"))
        state_font = fit(state, 68, 10.8, 7.5)
        c.setFillColor(palette["primary_dark"])
        c.setFont(bold_font, state_font)
        c.drawString(status_x + 9, page_h - margin - 50, txt(_pdf_clip_text(state, 68, bold_font, state_font)))
        c.setFont(regular_font, 6.7)
        c.setFillColor(palette["muted"])
        c.drawRightString(page_w - margin - 14, page_h - margin - header_h + 11, txt(f"Emitido em {printed_at}"))

        metric_y = page_h - margin - header_h - 54
        gap = 8
        chip_w = (page_w - (2 * margin) - (3 * gap)) / 4
        metrics = [
            ("Stock fisico", f"{self._fmt(qty)} {unit}", True),
            ("Disponivel", f"{self._fmt(available)} {unit}", False),
            ("Preco/unidade", money(price_unit), False),
            ("Valor em stock", money(stock_value), True),
        ]
        for index, (label, value, accent) in enumerate(metrics):
            chip(c, margin + index * (chip_w + gap), metric_y, chip_w, 42, label, value, accent)

        section_title(c, "Identificacao", metric_y - 34)
        info_y = metric_y - 72
        info_cols = [
            ("Categoria", category),
            ("Subcategoria", subcat),
            ("Tipo", kind),
            ("Dimensoes", dim_text),
            ("Fabricante", str(product.get("fabricante", "") or "-").strip() or "-"),
            ("Modelo", str(product.get("modelo", "") or "-").strip() or "-"),
            ("Peso/unid.", f"{self._fmt(product.get('peso_unid', 0))} kg"),
            ("Metros/unid.", self._fmt(product.get("metros_unidade", product.get("metros", 0)))),
            ("Compra", money(product.get("p_compra", 0))),
            ("PVP1", money(product.get("pvp1", 0))),
            ("PVP2", money(product.get("pvp2", 0))),
            ("Atualizado", str(product.get("atualizado_em", product.get("updated_at", "")) or "-").replace("T", " ")[:19] or "-"),
        ]
        box_w = (page_w - (2 * margin) - (2 * gap)) / 3
        box_h = 32
        for index, (label, value) in enumerate(info_cols):
            row = index // 3
            col = index % 3
            x = margin + col * (box_w + gap)
            y = info_y - row * (box_h + 7)
            chip(c, x, y, box_w, box_h, label, str(value), False)

        obs = str(product.get("obs", "") or "").strip()
        obs_y = info_y - 4 * (box_h + 7) - 10
        section_title(c, "Observacoes", obs_y + 30)
        c.setFillColor(palette["surface"])
        c.setStrokeColor(palette["line"])
        c.roundRect(margin, obs_y - 20, page_w - (2 * margin), 44, 9, stroke=1, fill=1)
        c.setFillColor(palette["ink"])
        c.setFont(regular_font, 8.3)
        for line_index, line in enumerate(_pdf_wrap_text(obs or "Sem observacoes registadas.", regular_font, 8.3, page_w - (2 * margin) - 18, max_lines=3)):
            c.drawString(margin + 9, obs_y + 8 - (line_index * 10), txt(line))

        table_top = obs_y - 54
        section_title(c, "Movimentos recentes", table_top + 24)
        cols = [("Data", 94), ("Tipo", 92), ("Operador", 106), ("Qtd", 54), ("Antes", 54), ("Depois", 58), ("Obs.", page_w - (2 * margin) - 458)]
        x = margin
        c.setFillColor(palette["primary_soft"])
        c.setStrokeColor(palette["line"])
        c.roundRect(margin, table_top - 4, page_w - (2 * margin), 22, 7, stroke=1, fill=1)
        c.setFillColor(palette["primary_dark"])
        c.setFont(bold_font, 7.4)
        for label, width in cols:
            c.drawString(x + 5, table_top + 3, txt(label))
            x += width
        y = table_top - 24
        if not movements:
            c.setFillColor(palette["muted"])
            c.setFont(regular_font, 8)
            c.drawString(margin + 5, y + 7, txt("Sem movimentos registados."))
        for index, mov in enumerate(movements):
            c.setFillColor(palette["surface"] if index % 2 == 0 else palette["surface_alt"])
            c.setStrokeColor(palette["line"])
            c.roundRect(margin, y, page_w - (2 * margin), 20, 5, stroke=1, fill=1)
            values = [
                str(mov.get("data", "") or "-")[:16],
                str(mov.get("tipo", "") or "-"),
                str(mov.get("operador", "") or "-"),
                self._fmt(mov.get("qtd", 0)),
                self._fmt(mov.get("antes", 0)),
                self._fmt(mov.get("depois", 0)),
                str(mov.get("obs", "") or "-").split("|meta|", 1)[0].strip() or "-",
            ]
            x = margin
            c.setFillColor(palette["ink"])
            c.setFont(regular_font, 7.1)
            for value, (_label, width) in zip(values, cols):
                c.drawString(x + 5, y + 7, txt(_pdf_clip_text(value, width - 10, regular_font, 7.1)))
                x += width
            y -= 21

        barcode_y = margin + 42
        c.setFillColor(palette["surface"])
        c.setStrokeColor(palette["line"])
        c.roundRect(margin, barcode_y, page_w - (2 * margin), 46, 10, stroke=1, fill=1)
        c.setFillColor(palette["muted"])
        c.setFont(regular_font, 7)
        c.drawString(margin + 10, barcode_y + 32, txt("Codigo para picagem / consulta"))
        self._draw_code128_fit(c, scan_code, margin + 12, barcode_y + 12, page_w - (2 * margin) - 24, 16, min_bar_width=0.34, max_bar_width=0.86)
        c.setFillColor(palette["ink"])
        c.setFont(bold_font, 7.6)
        c.drawCentredString(page_w / 2, barcode_y + 4, txt(_pdf_clip_text(scan_code, page_w - (2 * margin) - 24, bold_font, 7.6)))
        c.setFillColor(palette["muted"])
        c.setFont(regular_font, 7)
        c.drawRightString(page_w - margin, margin + 14, txt("LUGEST | Ficha de produto"))
        c.save()
        return target

    def product_open_sheet_pdf(self, codigo: str) -> Path:
        target = self.product_sheet_pdf(codigo)
        os.startfile(str(target))
        return target

    def _draw_product_stock_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        product: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        code = str(product.get("codigo", "") or "-").strip() or "-"
        scan_code = str(product.get("scan_code", "") or self.inventory_scan_code("PRD", code)).strip()
        unit = str(product.get("unid", "") or "UN").strip() or "UN"
        _draw_light_product_label(
            canvas_obj,
            page_width,
            page_height,
            product,
            palette,
            logo_path,
            printed_at,
            draw_logo=self._draw_operator_logo_plate,
            draw_barcode=self._draw_code128_fit,
            scan_code=scan_code,
            quantity_text=f"{self._fmt(product.get('qty', 0))} {unit}",
            price_text=f"{self._fmt(product.get('preco_unid', 0))} EUR",
            dimension_text=self._product_dimensoes(product),
        )
        return

    def product_label_pdf(self, codigo: str, output_path: str | Path | None = None) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        product = self.product_detail(codigo)
        code = str(product.get("codigo", "") or "").strip() or "produto"
        target = (
            Path(output_path)
            if output_path
            else self._storage_output_path("products/labels", f"Etiqueta_Produto_{code}.pdf")
        )
        target.parent.mkdir(parents=True, exist_ok=True)
        page_size = (110 * mm, 50 * mm)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=page_size)
        with pdf_font_size_limit(14.0):
            self._draw_product_stock_label(canvas_obj, page_size[0], page_size[1], product, palette, logo_path, printed_at)
        canvas_obj.save()
        return target

    def product_open_label_pdf(self, codigo: str) -> Path:
        target = self.product_label_pdf(codigo)
        os.startfile(str(target))
        return target
