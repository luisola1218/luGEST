"""Route document rendering, independent of the Qt bridge and live records."""
from dataclasses import dataclass
from typing import Any, Callable
from pathlib import Path
from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text

@dataclass(frozen=True)
class RouteReportRules:
    palette: Callable
    branding: Callable
    now_iso: Callable
    text: Callable
    draw_logo: Callable
    number: Callable
    currency: Callable

class RouteReport:
    def __init__(self, rules: RouteReportRules):
        self.rules = rules

    def render(self, detail: dict, path: str | Path) -> Path:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas

        out_path = Path(path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page_w, page_h = A4
        margin = 32
        inner_w = page_w - (2 * margin)
        palette = self.rules.palette()
        branding = self.rules.branding()
        logo_text = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_text) if logo_text and Path(logo_text).exists() else None
        printed_at = str(self.rules.now_iso() or "").replace("T", " ")[:19]
        regular = "Helvetica"
        bold = "Helvetica-Bold"
        c = canvas.Canvas(str(out_path), pagesize=A4)
        c.setTitle(self.rules.text(f"Documento de viagem {detail.get('numero', '')}"))
        page_number = 1

        def text(value: Any) -> str:
            return self.rules.text(value)

        def draw_card(x: float, y: float, width: float, height: float, label: str, value: str, *, accent: bool = False) -> None:
            c.setFillColor(palette["primary_soft"] if accent else palette["surface"])
            c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
            c.roundRect(x, y, width, height, 3, stroke=1, fill=1)
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.2)
            c.drawString(x + 8, y + height - 11, text(_pdf_clip_text(label, width - 16, regular, 6.2)))
            value_font = 9.4
            while value_font > 6.5 and len(str(value or "")) * value_font * 0.52 > width - 16:
                value_font -= 0.3
            c.setFillColor(palette["ink"])
            c.setFont(bold, value_font)
            c.drawString(x + 8, y + 8, text(_pdf_clip_text(value or "-", width - 16, bold, value_font)))

        def draw_footer() -> None:
            c.setStrokeColor(palette["line"])
            c.line(margin, 27, page_w - margin, 27)
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.3)
            c.drawString(margin, 17, text(f"LUGEST | Documento de viagem {detail.get('numero', '-') or '-'}"))
            c.drawRightString(page_w - margin, 17, text(f"Pagina {page_number} | {printed_at}"))

        def draw_header() -> float:
            top = page_h - 26
            c.setFillColor(palette["surface"])
            c.rect(0, 0, page_w, page_h, stroke=0, fill=1)
            c.setFillColor(palette["primary"])
            c.rect(0, page_h - 9, page_w, 9, stroke=0, fill=1)
            if logo_path:
                self.rules.draw_logo(c, palette, logo_path, margin, top - 54, 82, 40, radius=3, padding_x=4, padding_y=3)
            title_x = margin + 96
            title_w = inner_w - 190
            c.setFillColor(palette["primary_dark"])
            c.setFont(bold, 17)
            c.drawString(title_x, top - 21, text("Transportes | Documento de viagem"))
            c.setFillColor(palette["muted"])
            c.setFont(regular, 7.4)
            c.drawString(title_x, top - 37, text(_pdf_clip_text(f"Viagem {detail.get('numero', '-') or '-'} | {detail.get('data_planeada', '-') or '-'} as {detail.get('hora_saida', '-') or '-'}", title_w, regular, 7.4)))
            status = str(detail.get("estado", "") or "Planeado")
            draw_card(page_w - margin - 86, top - 52, 86, 38, "Estado", status, accent=True)

            metric_y = top - 101
            gap = 7
            metric_w = (inner_w - (3 * gap)) / 4
            metrics = [
                ("Paragens", str(len(list(detail.get("paragens", []) or [])))),
                ("Paletes", self.rules.number(detail.get("paletes", 0))),
                ("Peso bruto", f"{self.rules.number(detail.get('peso_bruto_kg', 0))} kg"),
                ("Volume", f"{self.rules.number(detail.get('volume_m3', 0))} m3"),
            ]
            for index, (label, value) in enumerate(metrics):
                draw_card(margin + index * (metric_w + gap), metric_y, metric_w, 35, label, value, accent=index == 0)

            meta_y = metric_y - 66
            meta_gap = 8
            meta_w = (inner_w - meta_gap) / 2
            vehicle = " | ".join(
                part for part in (
                    str(detail.get("viatura", "") or "").strip(),
                    str(detail.get("matricula", "") or "").strip(),
                    str(detail.get("motorista", "") or "").strip(),
                    str(detail.get("telefone_motorista", "") or "").strip(),
                ) if part
            ) or "Por definir"
            carrier = " | ".join(
                part for part in (
                    str(detail.get("transportadora_nome", "") or "").strip(),
                    str(detail.get("referencia_transporte", "") or "").strip(),
                    str(detail.get("pedido_transporte_estado", "") or "").strip(),
                    str(detail.get("pedido_transporte_ref", "") or "").strip(),
                ) if part
            ) or "Transporte proprio / sem pedido externo"
            draw_card(margin, meta_y, meta_w, 50, "Viatura / matricula / motorista / contacto", vehicle)
            draw_card(margin + meta_w + meta_gap, meta_y, meta_w, 50, "Transportadora / referencia / pedido", carrier)
            return meta_y - 16

        columns = [
            ("Ord", 28), ("Encomenda", 74), ("Cliente", 90), ("Descarga", 138),
            ("Planeado", 72), ("Guia", 58), ("Estado", inner_w - 460),
        ]

        def draw_table_header(y: float) -> float:
            c.setFillColor(palette["primary_soft"])
            c.setStrokeColor(palette["line_strong"])
            c.rect(margin, y - 20, inner_w, 20, stroke=1, fill=1)
            c.setFillColor(palette["primary_dark"])
            c.setFont(bold, 6.8)
            x = margin
            for label, width in columns:
                c.drawString(x + 5, y - 13, text(label))
                x += width
            return y - 25

        def start_new_page() -> float:
            nonlocal page_number
            draw_footer()
            c.showPage()
            page_number += 1
            return draw_table_header(draw_header())

        y = draw_table_header(draw_header())
        for row_index, stop in enumerate(list(detail.get("paragens", []) or [])):
            checklist = (
                f"Carga {'OK' if stop.get('check_carga_ok') else '-'} | "
                f"Docs {'OK' if stop.get('check_docs_ok') else '-'} | "
                f"Paletes {'OK' if stop.get('check_paletes_ok') else '-'}"
            )
            notes = " | ".join(
                part for part in (
                    f"Zona {stop.get('zona_transporte', '-') or '-'}",
                    (
                        f"GPS {str(stop.get('latitude', '') or '').strip()},"
                        f"{str(stop.get('longitude', '') or '').strip()}"
                        if str(stop.get("latitude", "") or "").strip()
                        and str(stop.get("longitude", "") or "").strip()
                        else ""
                    ),
                    f"Carga {self.rules.number(stop.get('paletes', 0))} pal / {self.rules.number(stop.get('peso_bruto_kg', 0))} kg / {self.rules.number(stop.get('volume_m3', 0))} m3",
                    checklist,
                    f"POD {stop.get('pod_estado', '-') or '-'}",
                    str(stop.get("observacoes", "") or "").strip(),
                ) if part
            )
            note_lines = _pdf_wrap_text(notes, regular, 6.4, inner_w - 18, max_lines=2) or []
            row_height = 27 + (len(note_lines) * 8)
            if y - row_height < 84:
                y = start_new_page()
            row_y = y - row_height
            c.setFillColor(palette["surface"] if row_index % 2 == 0 else palette["surface_alt"])
            c.setStrokeColor(palette["line"])
            c.rect(margin, row_y, inner_w, row_height - 3, stroke=1, fill=1)
            values = [
                str(stop.get("ordem", "-") or "-"),
                str(stop.get("encomenda_numero", "-") or "-"),
                str(stop.get("cliente_nome", "-") or "-"),
                str(stop.get("local_descarga", "-") or "-"),
                str(stop.get("data_planeada", "") or detail.get("data_planeada", "-")).replace("T", " ")[:16] or "-",
                str(stop.get("guia_numero", "-") or "-"),
                str(stop.get("estado", "-") or "-"),
            ]
            x = margin
            for column_index, (value, (_label, width)) in enumerate(zip(values, columns)):
                font_name = bold if column_index in (0, 1, 6) else regular
                c.setFillColor(palette["ink"] if column_index in (0, 1, 6) else palette["muted"])
                c.setFont(font_name, 6.7)
                c.drawString(x + 5, y - 17, text(_pdf_clip_text(value, width - 10, font_name, 6.7)))
                x += width
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.4)
            note_y = y - 27
            for line in note_lines:
                c.drawString(margin + 9, note_y, text(line))
                note_y -= 8
            y = row_y - 5

        if y < 92:
            y = start_new_page()
        signature_y = 50
        c.setStrokeColor(palette["line_strong"])
        signature_w = (inner_w - 20) / 3
        for index, label in enumerate(("Motorista / saida", "Conferencia de carga", "Rececao / chegada")):
            x = margin + index * (signature_w + 10)
            c.line(x, signature_y + 16, x + signature_w, signature_y + 16)
            c.setFillColor(palette["muted"])
            c.setFont(regular, 6.2)
            c.drawCentredString(x + signature_w / 2, signature_y + 6, text(label))
        draw_footer()
        c.save()
        return out_path


    def _transport_route_sheetrender_legacy(self, detail: dict, path: str | Path) -> Path:
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
                f"{detail.get('volume_m3', 0):.3f} m3 | Preco {self.rules.currency(detail.get('preco_total', 0))} | "
                f"Custo {self.rules.currency(detail.get('custo_total', 0))} | Sug. {self.rules.currency(detail.get('custo_sugerido_total', 0))}",
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
                f"Pal {self.rules.number(stop.get('paletes', 0))} | "
                f"{self.rules.number(stop.get('peso_bruto_kg', 0))} kg | "
                f"{self.rules.number(stop.get('volume_m3', 0))} m3 | "
                f"Preco {self.rules.currency(stop.get('preco_transporte', 0))} | "
                f"Custo {self.rules.currency(stop.get('custo_transporte', 0))} | Sug. {self.rules.currency(stop.get('custo_sugerido', 0))}"
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
