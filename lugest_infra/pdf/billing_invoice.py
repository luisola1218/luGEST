from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any


def _fmt_money(value: Any) -> str:
    try:
        return f"{float(value or 0):.2f} EUR"
    except Exception:
        return "0.00 EUR"


def _fmt_num(value: Any) -> str:
    try:
        val = float(value or 0)
    except Exception:
        val = 0.0
    if abs(val - round(val)) < 1e-9:
        return str(int(round(val)))
    return f"{val:.2f}"


def _fmt_date(value: Any) -> str:
    raw = str(value or "").strip().replace("T", " ")
    if not raw:
        return "-"
    try:
        return datetime.fromisoformat(raw).strftime("%d/%m/%Y")
    except Exception:
        if len(raw) >= 10 and raw[4:5] == "-" and raw[7:8] == "-":
            return f"{raw[8:10]}/{raw[5:7]}/{raw[:4]}"
        return raw[:10]


def _clip_text(value: Any, max_width: float, font_name: str, font_size: float) -> str:
    from reportlab.pdfbase import pdfmetrics

    text = str(value or "")
    if pdfmetrics.stringWidth(text, font_name, font_size) <= max_width:
        return text
    ellipsis = "..."
    while text and pdfmetrics.stringWidth(text + ellipsis, font_name, font_size) > max_width:
        text = text[:-1]
    return f"{text}{ellipsis}" if text else ""


def _wrap_text(value: Any, font_name: str, font_size: float, max_width: float, max_lines: int | None = None) -> list[str]:
    from reportlab.pdfbase import pdfmetrics

    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    if not text:
        return []
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip() if current else word
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
        if max_lines and len(lines) >= max_lines:
            break
    if current and (not max_lines or len(lines) < max_lines):
        lines.append(current)
    return lines[: max_lines or len(lines)]


def _pdf_palette(desktop_main) -> dict[str, Any]:
    from reportlab.lib import colors

    primary_hex = str((desktop_main.get_branding_config() or {}).get("primary_color", "#0F1F5C") or "#0F1F5C")
    mix_fn = getattr(desktop_main, "_orc_pdf_mix_hex", None)
    if not callable(mix_fn):
        mix_fn = getattr(desktop_main, "_pdf_mix_hex", None)
    if not callable(mix_fn):
        mix_fn = lambda base, target, _ratio: target if target else base
    return {
        "primary": colors.HexColor(primary_hex),
        "primary_dark": colors.HexColor(mix_fn(primary_hex, "#000000", 0.18)),
        "primary_soft": colors.HexColor(mix_fn(primary_hex, "#FFFFFF", 0.84)),
        "primary_soft_2": colors.HexColor(mix_fn(primary_hex, "#FFFFFF", 0.92)),
        "line": colors.HexColor(mix_fn(primary_hex, "#D9E2EC", 0.76)),
        "line_strong": colors.HexColor(mix_fn(primary_hex, "#6B7C93", 0.42)),
        "ink": colors.HexColor(mix_fn(primary_hex, "#111827", 0.72)),
        "muted": colors.HexColor("#52637A"),
        "surface_alt": colors.HexColor("#F8FAFC"),
    }


def _pdf_register_fonts() -> dict[str, str]:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    regular = "Helvetica"
    bold = "Helvetica-Bold"
    candidates = [
        ("Arial", r"C:\Windows\Fonts\arial.ttf"),
        ("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"),
        ("SegoeUI", r"C:\Windows\Fonts\segoeui.ttf"),
        ("SegoeUI-Bold", r"C:\Windows\Fonts\segoeuib.ttf"),
    ]
    for name, path in candidates:
        try:
            if path and Path(path).exists():
                pdfmetrics.registerFont(TTFont(name, path))
        except Exception:
            pass
    try:
        names = set(pdfmetrics.getRegisteredFontNames())
        if "Arial" in names:
            regular = "Arial"
        elif "SegoeUI" in names:
            regular = "SegoeUI"
        if "Arial-Bold" in names:
            bold = "Arial-Bold"
        elif "SegoeUI-Bold" in names:
            bold = "SegoeUI-Bold"
    except Exception:
        pass
    return {"regular": regular, "bold": bold}


def _draw_qr(canvas_obj, desktop_main, payload: str, atcud: str, x: float, top_y: float, size: float = 72) -> None:
    from reportlab.graphics import renderPDF
    from reportlab.graphics.barcode import qr as rl_qr
    from reportlab.graphics.shapes import Drawing

    ntxt = getattr(desktop_main, "pdf_normalize_text", lambda value: str(value or ""))
    page_height = 0.0
    try:
        page_height = float((getattr(canvas_obj, "_pagesize", (0, 0)) or (0, 0))[1] or 0.0)
    except Exception:
        page_height = 0.0
    draw_y = (page_height - top_y - size) if page_height > 0 else (top_y - size)
    try:
        widget = rl_qr.QrCodeWidget(payload)
        bounds = widget.getBounds()
        width = max(float(bounds[2] - bounds[0]), 1.0)
        height = max(float(bounds[3] - bounds[1]), 1.0)
        drawing = Drawing(size, size, transform=[size / width, 0, 0, size / height, 0, 0])
        drawing.add(widget)
        renderPDF.draw(drawing, canvas_obj, x, draw_y)
    except Exception:
        canvas_obj.setFont("Helvetica", 6.4)
        fallback_label = ntxt("QR indisponível")
        fallback_y = (draw_y + (size / 2.0)) if page_height > 0 else (top_y - 10)
        canvas_obj.drawString(x, fallback_y, fallback_label)


def render_invoice_pdf(backend, output_path: str | Path, doc: dict[str, Any]) -> Path:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas

    desktop_main = backend.desktop_main
    normalize_text = getattr(desktop_main, "pdf_normalize_text", lambda value: str(value or ""))

    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)

    fonts = _pdf_register_fonts()
    palette = _pdf_palette(desktop_main)
    width, height = A4
    margin = 24
    content_w = width - (margin * 2)
    row_h = 22
    footer_top = 786
    first_table_top = 306
    next_table_top = 92
    reserved_bottom = 232
    lines = list(doc.get("lines", []) or [])

    table_cols = [
        ("Ref.", 104, "w"),
        ("Descricao", 172, "w"),
        ("Qtd.", 40, "e"),
        ("Un.", 30, "center"),
        ("P. Unit.", 58, "e"),
        ("IVA", 32, "e"),
        ("Base", 54, "e"),
        ("Total", 57, "e"),
    ]
    table_w = sum(width_col for _, width_col, _ in table_cols)
    first_capacity = max(1, int((footer_top - first_table_top - reserved_bottom) // row_h))
    next_capacity = max(1, int((footer_top - next_table_top - reserved_bottom) // row_h))
    if len(lines) <= first_capacity:
        total_pages = 1
    else:
        remaining = len(lines) - first_capacity
        total_pages = 1 + ((remaining + next_capacity - 1) // next_capacity)

    canvas_obj = pdf_canvas.Canvas(str(target), pagesize=A4)
    canvas_obj.setTitle(str(doc.get("legal_invoice_no", "") or doc.get("numero_fatura", "") or "Fatura"))
    try:
        canvas_obj.setPageCompression(0)
    except Exception:
        pass

    def yinv(top_y: float) -> float:
        return height - top_y

    def ntxt(value: Any) -> str:
        return normalize_text(value)

    def chip(x: float, top_y: float, width_box: float, label: str, value: str, height_box: float = 24) -> None:
        canvas_obj.saveState()
        canvas_obj.setFillColor(colors.white)
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(x, yinv(top_y + height_box), width_box, height_box, 7, stroke=1, fill=1)
        canvas_obj.restoreState()
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(fonts["regular"], 6.8)
        canvas_obj.drawString(x + 7, yinv(top_y + 8.8), ntxt(label))
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(fonts["bold"], 9.2)
        canvas_obj.drawString(x + 7, yinv(top_y + 18.2), ntxt(_clip_text(value, width_box - 14, fonts["bold"], 9.2)))

    def block(x: float, top_y: float, width_box: float, title: str, rows: list[str], height_box: float = 96) -> None:
        canvas_obj.saveState()
        canvas_obj.setFillColor(colors.white)
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.roundRect(x, yinv(top_y + height_box), width_box, height_box, 8, stroke=1, fill=1)
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.roundRect(x + 1, yinv(top_y + 24), width_box - 2, 22, 8, stroke=0, fill=1)
        canvas_obj.restoreState()
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(fonts["bold"], 9.2)
        canvas_obj.drawString(x + 8, yinv(top_y + 14), ntxt(title))
        yy = top_y + 36
        for idx, row in enumerate(rows):
            font_name = fonts["bold"] if idx == 0 else fonts["regular"]
            font_size = 9.4 if idx == 0 else 8.6
            wrapped = _wrap_text(row, font_name, font_size, width_box - 16, max_lines=2)
            for item in wrapped:
                canvas_obj.setFillColor(palette["ink"])
                canvas_obj.setFont(font_name, font_size)
                canvas_obj.drawString(x + 8, yinv(yy), ntxt(item))
                yy += 10.8
                if yy > top_y + height_box - 8:
                    return

    def draw_table_header(top_y: float) -> float:
        canvas_obj.saveState()
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(0.8)
        canvas_obj.roundRect(margin, yinv(top_y + 22), table_w, 22, 7, stroke=1, fill=1)
        canvas_obj.restoreState()
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(fonts["bold"], 8.4)
        xx = margin
        for label, width_col, align in table_cols:
            if align == "e":
                canvas_obj.drawRightString(xx + width_col - 7, yinv(top_y + 14.2), ntxt(label))
            elif align == "center":
                canvas_obj.drawCentredString(xx + (width_col / 2.0), yinv(top_y + 14.2), ntxt(label))
            else:
                canvas_obj.drawString(xx + 7, yinv(top_y + 14.2), ntxt(label))
            xx += width_col
        return top_y + 26

    def draw_row(top_y: float, row_index: int, line: dict[str, Any]) -> None:
        fill = palette["surface_alt"] if row_index % 2 == 0 else colors.white
        canvas_obj.saveState()
        canvas_obj.setFillColor(fill)
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.roundRect(margin, yinv(top_y + row_h), table_w, row_h, 5, stroke=1, fill=1)
        canvas_obj.restoreState()
        values = [
            str(line.get("reference", "") or "-").strip(),
            str(line.get("description", "") or "-").strip(),
            _fmt_num(line.get("quantity", 0)),
            str(line.get("unit", "") or "UN").strip(),
            _fmt_money(line.get("unit_price", 0)).replace(" EUR", ""),
            f"{_fmt_num(line.get('iva_perc', 0))}%",
            _fmt_money(line.get("subtotal", 0)).replace(" EUR", ""),
            _fmt_money(line.get("total", 0)).replace(" EUR", ""),
        ]
        xx = margin
        for (label, width_col, align), value in zip(table_cols, values):
            canvas_obj.setFillColor(palette["ink"])
            font_name = fonts["bold"] if label == "Ref." else fonts["regular"]
            font_size = 8.3 if label == "Ref." else 8.2
            canvas_obj.setFont(font_name, font_size)
            clipped = _clip_text(value, width_col - 14, font_name, font_size)
            if label == "Descricao":
                clipped = _clip_text(value, width_col - 14, fonts["regular"], 7.8)
            if align == "e":
                canvas_obj.drawRightString(xx + width_col - 7, yinv(top_y + 13.4), ntxt(clipped))
            elif align == "center":
                canvas_obj.drawCentredString(xx + (width_col / 2.0), yinv(top_y + 13.4), ntxt(clipped))
            else:
                canvas_obj.drawString(xx + 7, yinv(top_y + 13.4), ntxt(clipped))
            xx += width_col

    def draw_footer(page_no: int) -> None:
        issuer = dict(doc.get("issuer", {}) or {})
        references = dict(doc.get("references", {}) or {})
        canvas_obj.saveState()
        canvas_obj.setStrokeColor(palette["line_strong"])
        canvas_obj.line(margin, yinv(footer_top), width - margin, yinv(footer_top))
        canvas_obj.restoreState()
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.setFont(fonts["regular"], 7.4)
        canvas_obj.drawString(
            margin,
            yinv(footer_top + 14),
            ntxt(
                (
                    f"Documento processado por programa certificado n.o {str(doc.get('software_cert', '') or '').strip()} | Pag. {page_no}/{total_pages}"
                    if str(doc.get("software_cert", "") or "").strip()
                    else f"Documento emitido por LuGEST | Pag. {page_no}/{total_pages}"
                )
            ),
        )
        canvas_obj.drawString(
            margin,
            yinv(footer_top + 26),
            ntxt(
                f"Origem venda: Orc. {references.get('orcamento', '-') or '-'} | Enc. {references.get('encomenda', '-') or '-'} | Guia {references.get('guia', '-') or '-'}"
            ),
        )
        obs = str(doc.get("obs", "") or "").strip()
        if obs:
            obs_line = _clip_text(f"Obs.: {obs}", content_w, fonts["regular"], 7.4)
            canvas_obj.drawString(margin, yinv(footer_top + 38), ntxt(obs_line))
        if issuer:
            canvas_obj.drawRightString(width - margin, yinv(footer_top + 14), ntxt(str(issuer.get("nome", "") or "").strip() or "Emitente"))
            canvas_obj.drawRightString(width - margin, yinv(footer_top + 26), ntxt(f"NIF {str(issuer.get('nif', '') or '-').strip()}"))

    def draw_first_page_header() -> float:
        banner_top = 24
        banner_h = 106
        inner_pad = 14
        logo_slot_w = 102
        logo_plate_gap = 12
        metrics_w = 206
        section_gap = 16
        banner_x = margin + logo_slot_w + logo_plate_gap
        banner_w = width - margin - banner_x
        draw_header_panel = getattr(desktop_main, "draw_pdf_header_panel", None)
        if callable(draw_header_panel):
            try:
                draw_header_panel(canvas_obj, height, banner_x, banner_top, banner_w, banner_h, radius=12, stroke_color="#D5DDE7", accent_color="#EAF0F6", accent_height=5)
            except Exception:
                draw_header_panel = None
        if not callable(draw_header_panel):
            canvas_obj.saveState()
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(palette["line_strong"])
            canvas_obj.setLineWidth(1.0)
            canvas_obj.roundRect(banner_x, yinv(banner_top + banner_h), banner_w, banner_h, 12, stroke=1, fill=1)
            canvas_obj.restoreState()

        draw_logo_plate = getattr(desktop_main, "draw_pdf_logo_plate", None)
        if callable(draw_logo_plate):
            try:
                draw_logo_plate(
                    canvas_obj,
                    height,
                    margin,
                    50,
                    box_w=logo_slot_w,
                    box_h=54,
                    padding=4,
                )
            except Exception:
                pass

        banner_inner_x = banner_x + inner_pad
        banner_inner_w = banner_w - (inner_pad * 2)
        metrics_x = banner_x + banner_w - inner_pad - metrics_w
        title_x = banner_inner_x + section_gap
        title_w = max(120.0, metrics_x - title_x - section_gap)

        canvas_obj.setFillColor(palette["primary_dark"])
        title_text = str(doc.get("titulo", "Fatura")) or "Fatura"
        subtitle_text = str(doc.get("subtitulo", "Documento comercial")) or "Documento comercial"
        canvas_obj.setFont(fonts["bold"], 19)
        canvas_obj.drawCentredString(title_x + (title_w / 2.0), yinv(50), ntxt(title_text))
        canvas_obj.setFont(fonts["regular"], 9.6)
        canvas_obj.setFillColor(palette["muted"])
        canvas_obj.drawCentredString(
            title_x + (title_w / 2.0),
            yinv(68),
            ntxt(_clip_text(subtitle_text, title_w, fonts["regular"], 9.6)),
        )

        gap = 6
        chip_w = (metrics_w - gap) / 2.0
        chip(metrics_x, 32, chip_w, "Documento", str(doc.get("legal_invoice_no", "") or doc.get("numero_fatura", "") or "-"))
        chip(metrics_x + chip_w + gap, 32, chip_w, "Série", str(doc.get("serie", "") or "-"))
        chip(metrics_x, 58, chip_w, "Emissão", _fmt_date(doc.get("data_emissao", "")))
        chip(metrics_x + chip_w + gap, 58, chip_w, "Vencimento", _fmt_date(doc.get("data_vencimento", "")))
        chip(metrics_x, 84, chip_w, "Moeda", str(doc.get("moeda", "EUR") or "EUR"))
        chip(
            metrics_x + chip_w + gap,
            84,
            chip_w,
            "Guia",
            str((doc.get("references", {}) or {}).get("guia", "") or "-"),
        )

        issuer = dict(doc.get("issuer", {}) or {})
        customer = dict(doc.get("customer", {}) or {})
        top_cards_gap = 12
        top_card_w = (content_w - top_cards_gap) / 2.0
        top_card_h = 92
        block(
            margin,
            124,
            top_card_w,
            "Emitente",
            [
                str(issuer.get("nome", "") or "-").strip(),
                f"NIF: {str(issuer.get('nif', '') or '-').strip()}",
                str(issuer.get("morada", "") or "-").strip(),
                str(issuer.get("extra", "") or "").strip(),
            ],
            height_box=top_card_h,
        )
        block(
            margin + top_card_w + top_cards_gap,
            124,
            top_card_w,
            "Cliente",
            [
                str(customer.get("nome", "") or "-").strip(),
                f"NIF: {str(customer.get('nif', '') or '-').strip()}",
                str(customer.get("morada", "") or "-").strip(),
                f"Contacto: {str(customer.get('contacto', '') or '-').strip()}",
            ],
            height_box=top_card_h,
        )
        refs = dict(doc.get("references", {}) or {})
        refs_txt = (
            f"Orçamento: {str(refs.get('orcamento', '') or '-').strip()} | "
            f"Encomenda: {str(refs.get('encomenda', '') or '-').strip()} | "
            f"Guia: {str(refs.get('guia', '') or '-').strip()}"
        )
        block(margin, 228, content_w, "Referências", [refs_txt], height_box=50)
        return draw_table_header(first_table_top)

    def draw_next_page_header(page_no: int) -> float:
        canvas_obj.saveState()
        canvas_obj.setFillColor(palette["primary_soft"])
        canvas_obj.roundRect(margin, yinv(24 + 48), content_w, 48, 12, stroke=0, fill=1)
        canvas_obj.restoreState()
        canvas_obj.setFillColor(palette["primary_dark"])
        canvas_obj.setFont(fonts["bold"], 15)
        canvas_obj.drawString(margin + 12, yinv(43), ntxt(str(doc.get("legal_invoice_no", "") or doc.get("numero_fatura", "") or "Fatura")))
        canvas_obj.setFont(fonts["regular"], 8.8)
        canvas_obj.drawString(margin + 12, yinv(58), ntxt(f"Cliente: {str((doc.get('customer', {}) or {}).get('nome', '') or '-').strip()}"))
        canvas_obj.drawRightString(width - margin, yinv(43), ntxt(f"Pag. {page_no}/{total_pages}"))
        return draw_table_header(next_table_top)

    remaining_lines = list(lines)
    for page_no in range(1, total_pages + 1):
        first_page = page_no == 1
        table_y = draw_first_page_header() if first_page else draw_next_page_header(page_no)
        capacity = first_capacity if first_page else next_capacity
        page_lines = remaining_lines[:capacity]
        remaining_lines = remaining_lines[capacity:]
        row_top = table_y
        for local_index, line in enumerate(page_lines):
            draw_row(row_top, local_index, line)
            row_top += row_h

        if page_no == total_pages:
            atcud_h = 98
            summary_h = 118
            summary_y = footer_top - summary_h - 6
            atcud_y = summary_y - atcud_h - 10

            summary_gap = 10
            fiscal_x = margin
            fiscal_w = (content_w - summary_gap) / 2.0
            canvas_obj.saveState()
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(palette["line_strong"])
            canvas_obj.roundRect(fiscal_x, yinv(summary_y + summary_h), fiscal_w, summary_h, 10, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["primary_soft"])
            canvas_obj.roundRect(fiscal_x + 1, yinv(summary_y + 24), fiscal_w - 2, 22, 9, stroke=0, fill=1)
            canvas_obj.restoreState()
            canvas_obj.setFillColor(palette["primary_dark"])
            canvas_obj.setFont(fonts["bold"], 8.8)
            canvas_obj.drawString(fiscal_x + 8, yinv(summary_y + 14), ntxt("Fiscalidade e documento"))
            tax_y = summary_y + 40
            for row in list(doc.get("tax_summary", []) or [])[:3]:
                canvas_obj.setFillColor(palette["ink"])
                canvas_obj.setFont(fonts["regular"], 8.2)
                canvas_obj.drawString(fiscal_x + 10, yinv(tax_y), ntxt(f"{str(row.get('label', '') or '').strip()}: {_fmt_money(row.get('base', 0))}"))
                tax_y += 12
                canvas_obj.drawString(fiscal_x + 10, yinv(tax_y), ntxt(f"IVA {str(row.get('rate_label', '') or '').strip()}: {_fmt_money(row.get('tax', 0))}"))
                tax_y += 12
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(fonts["regular"], 8.0)
            canvas_obj.drawString(
                fiscal_x + 10,
                yinv(summary_y + 94),
                ntxt(
                    f"Guia: {str((doc.get('references', {}) or {}).get('guia', '') or '-').strip()} | "
                    f"Vencimento: {_fmt_date(doc.get('data_vencimento', ''))}"
                ),
            )

            totals_x = margin + fiscal_w + summary_gap
            totals_w = content_w - fiscal_w - summary_gap
            canvas_obj.saveState()
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(palette["line_strong"])
            canvas_obj.roundRect(totals_x, yinv(atcud_y + atcud_h), totals_w, atcud_h, 10, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["primary_soft"])
            canvas_obj.roundRect(totals_x + 1, yinv(atcud_y + 24), totals_w - 2, 22, 9, stroke=0, fill=1)
            canvas_obj.restoreState()
            canvas_obj.setFillColor(palette["primary_dark"])
            canvas_obj.setFont(fonts["bold"], 8.8)
            canvas_obj.drawString(totals_x + 8, yinv(atcud_y + 14), ntxt("ATCUD e validação"))
            card_inner_pad = 12
            qr_frame_pad = 4
            body_top = atcud_y + 28
            body_bottom = atcud_y + atcud_h - card_inner_pad
            body_h = max(36.0, body_bottom - body_top)
            qr_frame_size = min(58.0, body_h)
            qr_frame_y = body_top + ((body_h - qr_frame_size) / 2.0)
            qr_frame_x = totals_x + totals_w - qr_frame_size - card_inner_pad
            qr_size = qr_frame_size - (qr_frame_pad * 2.0)
            qr_x = qr_frame_x + qr_frame_pad
            qr_top = qr_frame_y + qr_frame_pad
            text_w = max(90.0, (qr_frame_x - totals_x) - (card_inner_pad + 2))
            canvas_obj.saveState()
            canvas_obj.setFillColor(palette["primary_soft_2"])
            canvas_obj.setStrokeColor(palette["line"])
            canvas_obj.roundRect(qr_frame_x, yinv(qr_frame_y + qr_frame_size), qr_frame_size, qr_frame_size, 8, stroke=1, fill=1)
            canvas_obj.restoreState()
            canvas_obj.setFillColor(palette["ink"])
            canvas_obj.setFont(fonts["bold"], 9.4)
            canvas_obj.drawString(
                totals_x + 10,
                yinv(atcud_y + 38),
                ntxt(_clip_text(str(doc.get("atcud", "") or "Pendente"), text_w, fonts["bold"], 9.4)),
            )
            canvas_obj.setFillColor(palette["muted"])
            canvas_obj.setFont(fonts["regular"], 7.8)
            canvas_obj.drawString(
                totals_x + 10,
                yinv(atcud_y + 52),
                ntxt(_clip_text("Código único do documento fiscal", text_w, fonts["regular"], 7.8)),
            )
            canvas_obj.drawString(
                totals_x + 10,
                yinv(atcud_y + 66),
                ntxt(_clip_text(f"Doc.: {str(doc.get('numero_fatura', '') or '-').strip()}", text_w, fonts["regular"], 7.8)),
            )
            _draw_qr(
                canvas_obj,
                desktop_main,
                str(doc.get("qr_payload", "") or "").strip(),
                str(doc.get("atcud", "") or ""),
                qr_x,
                qr_top,
                size=qr_size,
            )

            canvas_obj.saveState()
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(palette["line_strong"])
            canvas_obj.roundRect(totals_x, yinv(summary_y + summary_h), totals_w, summary_h, 10, stroke=1, fill=1)
            canvas_obj.setFillColor(palette["primary_soft"])
            canvas_obj.roundRect(totals_x + 1, yinv(summary_y + 24), totals_w - 2, 22, 9, stroke=0, fill=1)
            canvas_obj.restoreState()
            canvas_obj.setFillColor(palette["primary_dark"])
            canvas_obj.setFont(fonts["bold"], 8.8)
            canvas_obj.drawString(totals_x + 8, yinv(summary_y + 14), ntxt("Totais da fatura"))
            totals = [
                ("Base tributavel", _fmt_money(doc.get("subtotal", 0))),
                ("IVA", _fmt_money(doc.get("valor_iva", 0))),
                ("Total", _fmt_money(doc.get("valor_total", 0))),
                ("Recebido", _fmt_money(doc.get("valor_recebido", 0))),
                ("Por regularizar", _fmt_money(doc.get("saldo", doc.get("valor_total", 0)))),
            ]
            yy = summary_y + 40
            for label, value in totals:
                canvas_obj.setFillColor(palette["muted"])
                canvas_obj.setFont(fonts["regular"], 8.5)
                canvas_obj.drawString(totals_x + 10, yinv(yy), ntxt(label))
                canvas_obj.setFillColor(palette["ink"])
                canvas_obj.setFont(fonts["bold"], 9.1)
                canvas_obj.drawRightString(totals_x + totals_w - 10, yinv(yy), ntxt(value))
                yy += 15

        draw_footer(page_no)
        if page_no < total_pages:
            canvas_obj.showPage()

    canvas_obj.save()
    return target
