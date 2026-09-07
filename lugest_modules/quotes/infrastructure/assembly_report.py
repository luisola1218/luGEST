from __future__ import annotations


import math
from pathlib import Path
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text


from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class AssemblyReportPorts:
    _draw_code128_fit: Callable[..., Any]
    _draw_operator_logo_plate: Callable[..., Any]
    _fmt: Callable[..., Any]
    _operator_label_palette: Callable[..., Any]
    _operator_pdf_text: Callable[..., Any]
    _storage_output_path: Callable[..., Any]
    branding_settings: Callable[..., Any]
    conjunto_detail: Callable[..., Any]
    now_iso: Callable[..., Any]
    orc_line_is_piece: Callable[..., Any]
    orc_line_is_product: Callable[..., Any]
    orc_line_is_service: Callable[..., Any]

def render_assembly_sheet(ports: AssemblyReportPorts, codigo: str, output_path: str | Path | None = None) -> Path:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as pdf_canvas

    detail = ports.conjunto_detail(codigo)
    code = str(detail.get("codigo", "") or "CONJUNTO").strip() or "CONJUNTO"
    param_code = str(detail.get("param_codigo", "") or "-").strip() or "-"
    technical = dict(detail.get("ficha_tecnica", {}) or {})
    items = list(detail.get("itens", []) or [])
    name = str(detail.get("descricao", "") or code).strip() or code
    safe_code = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in code).strip("_") or "conjunto"
    target = Path(output_path) if output_path else ports._storage_output_path("quotes/assembly_sheets", f"Dossier_Tecnico_{safe_code}.pdf")
    target.parent.mkdir(parents=True, exist_ok=True)

    page_w, page_h = A4
    margin = 11 * mm
    inner_w = page_w - 2 * margin
    navy = colors.HexColor("#0B1F33")
    steel = colors.HexColor("#34495E")
    cyan = ports._operator_label_palette()["primary"]
    green = colors.HexColor("#78BE20")
    ink = colors.HexColor("#14212B")
    muted = colors.HexColor("#61717F")
    line = colors.HexColor("#CAD3DA")
    surface = colors.HexColor("#F3F6F8")
    white = colors.white
    palette = ports._operator_label_palette()
    branding = ports.branding_settings()
    logo_txt = str(branding.get("logo_path", "") or "").strip()
    logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
    issued_at = str(ports.now_iso() or "").replace("T", " ")[:19]
    revision = str(technical.get("modelo_versao", "") or "REV. 00").strip() or "REV. 00"

    fabricated_count = sum(1 for item in items if ports.orc_line_is_piece(item) and not str(item.get("stock_material_id", "") or "").strip())
    stock_count = sum(1 for item in items if ports.orc_line_is_product(item) or str(item.get("stock_material_id", "") or "").strip())
    service_count = sum(1 for item in items if ports.orc_line_is_service(item))
    material_names = {
        str(item.get("material", "") or "").strip()
        for item in items
        if str(item.get("material", "") or "").strip()
    }
    rows_per_page = 20
    bom_pages = max(1, math.ceil(max(1, len(items)) / rows_per_page))
    total_pages = 1 + bom_pages
    canvas_obj = pdf_canvas.Canvas(str(target), pagesize=A4)
    canvas_obj.setTitle(ports._operator_pdf_text(f"Dossier tecnico {code}"))
    canvas_obj.setAuthor("LUGEST")
    canvas_obj.setSubject("Ficha tecnica e lista de materiais do conjunto")

    def txt(value: object) -> str:
        return ports._operator_pdf_text(value)

    def controlled_footer(page_number: int, section: str) -> None:
        canvas_obj.setStrokeColor(line)
        canvas_obj.line(margin, margin + 5.2 * mm, page_w - margin, margin + 5.2 * mm)
        canvas_obj.setFillColor(muted)
        canvas_obj.setFont("Helvetica", 5.8)
        canvas_obj.drawString(margin, margin + 1.8 * mm, txt(f"DOCUMENTO CONTROLADO | PARAM {param_code} | {revision}"))
        canvas_obj.drawCentredString(page_w / 2, margin + 1.8 * mm, txt(section))
        canvas_obj.drawRightString(page_w - margin, margin + 1.8 * mm, txt(f"{page_number}/{total_pages} | {issued_at}"))

    def label_value(x: float, y_top: float, width: float, label: str, value: str, height: float = 17 * mm) -> None:
        canvas_obj.setFillColor(surface)
        canvas_obj.setStrokeColor(line)
        canvas_obj.roundRect(x, y_top - height, width, height, 2.5, stroke=1, fill=1)
        canvas_obj.setFillColor(steel)
        canvas_obj.setFont("Helvetica-Bold", 6.1)
        canvas_obj.drawString(x + 3 * mm, y_top - 4.8 * mm, txt(label.upper()))
        canvas_obj.setFillColor(ink)
        canvas_obj.setFont("Helvetica", 7.4)
        height_mm = height / mm
        max_lines = max(1, min(4, int((height_mm - 7.0) // 3.5)))
        wrapped = _pdf_wrap_text(
            str(value or "POR DEFINIR") or "POR DEFINIR",
            "Helvetica",
            7.4,
            width - 6 * mm,
            max_lines=max_lines,
        ) or ["POR DEFINIR"]
        y_line = y_top - 9.3 * mm
        for wrapped_line in wrapped:
            canvas_obj.drawString(x + 3 * mm, y_line, txt(wrapped_line))
            y_line -= 3.5 * mm

    def section_title(y: float, index: str, title: str, subtitle: str = "") -> float:
        canvas_obj.setFillColor(surface)
        canvas_obj.setStrokeColor(line)
        canvas_obj.rect(margin, y - 8 * mm, inner_w, 8 * mm, stroke=1, fill=1)
        canvas_obj.setFillColor(cyan)
        canvas_obj.rect(margin, y - 8 * mm, 13 * mm, 8 * mm, stroke=0, fill=1)
        canvas_obj.setFillColor(white)
        canvas_obj.setFont("Helvetica-Bold", 7.2)
        canvas_obj.drawCentredString(margin + 6.5 * mm, y - 5.4 * mm, txt(index))
        canvas_obj.setFillColor(navy)
        canvas_obj.setFont("Helvetica-Bold", 8)
        canvas_obj.drawString(margin + 17 * mm, y - 5.4 * mm, txt(title.upper()))
        if subtitle:
            canvas_obj.setFillColor(muted)
            canvas_obj.setFont("Helvetica", 5.8)
            canvas_obj.drawRightString(page_w - margin - 3 * mm, y - 5.2 * mm, txt(subtitle))
        return y - 11 * mm

    # Capa tecnica: documento de engenharia, sem dados comerciais.
    canvas_obj.setFillColor(white)
    canvas_obj.rect(0, page_h - 51 * mm, page_w, 51 * mm, stroke=0, fill=1)
    canvas_obj.setStrokeColor(line)
    canvas_obj.line(margin, page_h - 51 * mm, page_w - margin, page_h - 51 * mm)
    canvas_obj.setFillColor(cyan)
    canvas_obj.rect(0, page_h - 3 * mm, page_w, 3 * mm, stroke=0, fill=1)
    ports._draw_operator_logo_plate(canvas_obj, palette, logo_path, margin, page_h - 34 * mm, 39 * mm, 17 * mm, radius=3, padding_x=3, padding_y=2)
    canvas_obj.setFillColor(navy)
    canvas_obj.setFont("Helvetica-Bold", 7)
    canvas_obj.drawString(margin + 47 * mm, page_h - 15 * mm, txt("DOSSIER TECNICO DE CONJUNTO"))
    title_text = txt(name)
    title_width = 104 * mm
    title_size = 14.5
    title_lines = _pdf_wrap_text(title_text, "Helvetica-Bold", title_size, title_width, max_lines=3)
    while title_size > 11.0 and len(title_lines) > 2:
        title_size -= 0.5
        title_lines = _pdf_wrap_text(title_text, "Helvetica-Bold", title_size, title_width, max_lines=3)
    title_lines = title_lines[:2] or [title_text]
    canvas_obj.setFont("Helvetica-Bold", title_size)
    title_y = page_h - (21.5 * mm if len(title_lines) > 1 else 25 * mm)
    for title_line in title_lines:
        canvas_obj.drawString(margin + 47 * mm, title_y, txt(title_line))
        title_y -= 6 * mm
    canvas_obj.setFont("Helvetica", 7.2)
    canvas_obj.drawString(margin + 47 * mm, page_h - 39 * mm, txt(_pdf_clip_text(f"ENGENHARIA DE PRODUTO | {technical.get('familia_produto', '') or 'FAMILIA POR DEFINIR'}", 104 * mm, "Helvetica", 7.2)))
    canvas_obj.setFillColor(green)
    canvas_obj.roundRect(page_w - margin - 34 * mm, page_h - 18 * mm, 34 * mm, 8 * mm, 2, stroke=0, fill=1)
    canvas_obj.setFillColor(navy)
    canvas_obj.setFont("Helvetica-Bold", 6.2)
    canvas_obj.drawCentredString(page_w - margin - 17 * mm, page_h - 15.3 * mm, txt("LIBERTADO P/ PRODUCAO"))

    y = page_h - 61 * mm
    y = section_title(y, "01", "Identificacao e rastreabilidade", "Digital product passport")
    half = (inner_w - 4 * mm) / 2
    label_value(margin, y, half, "Codigo de produto", code, 15 * mm)
    label_value(margin + half + 4 * mm, y, half, "Codigo de parametrizacao", param_code, 15 * mm)
    y -= 18 * mm
    label_value(margin, y, half, "Modelo / revisao", revision, 15 * mm)
    label_value(margin + half + 4 * mm, y, half, "Aplicacao / destino", str(technical.get("aplicacao", "") or "POR DEFINIR"), 15 * mm)
    barcode_y = y - 33 * mm
    canvas_obj.setFillColor(white)
    canvas_obj.setStrokeColor(line)
    canvas_obj.roundRect(margin, barcode_y, inner_w, 14 * mm, 3, stroke=1, fill=1)
    canvas_obj.setFillColor(steel)
    canvas_obj.setFont("Helvetica-Bold", 5.7)
    canvas_obj.drawString(margin + 3 * mm, barcode_y + 9.2 * mm, txt("IDENTIFICADOR DIGITAL"))
    canvas_obj.setFont("Helvetica", 5.3)
    canvas_obj.drawString(margin + 3 * mm, barcode_y + 4.6 * mm, txt(_pdf_clip_text(f"PARAM {param_code} | {code}", 60 * mm, "Helvetica", 5.3)))
    barcode_value = f"CFG|{param_code}|{code}"
    ports._draw_code128_fit(canvas_obj, barcode_value, margin + 67 * mm, barcode_y + 3.7 * mm, inner_w - 72 * mm, 8 * mm, min_bar_width=0.25, max_bar_width=0.62)
    canvas_obj.setFont("Helvetica-Bold", 4.8)
    canvas_obj.drawCentredString(margin + 67 * mm + (inner_w - 72 * mm) / 2, barcode_y + 1.4 * mm, txt(_pdf_clip_text(barcode_value, inner_w - 72 * mm, "Helvetica-Bold", 4.8)))

    y = barcode_y - 5 * mm
    y = section_title(y, "02", "Definicao tecnica", "Especificacao funcional")
    label_value(margin, y, half, "Configuracao principal", str(technical.get("configuracao", "") or "POR DEFINIR"), 15 * mm)
    label_value(margin + half + 4 * mm, y, half, "Dimensoes / capacidade", str(technical.get("dimensoes_gerais", "") or "POR DEFINIR"), 15 * mm)
    y -= 18 * mm
    label_value(margin, y, half, "Materiais / acabamentos", str(technical.get("materiais_acabamentos", "") or "POR DEFINIR"), 15 * mm)
    label_value(margin + half + 4 * mm, y, half, "Instalacao / interfaces", str(technical.get("requisitos_instalacao", "") or "POR DEFINIR"), 15 * mm)
    y -= 18 * mm
    label_value(margin, y, inner_w, "Caracteristicas e desempenho", str(technical.get("caracteristicas", "") or "POR DEFINIR"), 17 * mm)

    y -= 21 * mm
    y = section_title(y, "03", "Composicao do produto", "Resumo da estrutura")
    metric_gap = 3 * mm
    metric_w = (inner_w - 4 * metric_gap) / 5
    metrics = [
        ("REFERENCIAS", len(items), cyan),
        ("FABRICADAS", fabricated_count, green),
        ("STOCK / MP", stock_count, colors.HexColor("#F0A202")),
        ("OPERACOES", service_count, colors.HexColor("#7A6FF0")),
        ("MATERIAIS", len(material_names), steel),
    ]
    for index, (label, value, accent) in enumerate(metrics):
        x = margin + index * (metric_w + metric_gap)
        canvas_obj.setFillColor(white)
        canvas_obj.setStrokeColor(line)
        canvas_obj.roundRect(x, y - 15 * mm, metric_w, 15 * mm, 3, stroke=1, fill=1)
        canvas_obj.setFillColor(accent)
        canvas_obj.rect(x, y - 15 * mm, 3 * mm, 15 * mm, stroke=0, fill=1)
        canvas_obj.setFillColor(muted)
        canvas_obj.setFont("Helvetica-Bold", 5.8)
        canvas_obj.drawString(x + 5 * mm, y - 5.3 * mm, txt(label))
        canvas_obj.setFillColor(ink)
        canvas_obj.setFont("Helvetica-Bold", 14)
        canvas_obj.drawString(x + 5 * mm, y - 11.8 * mm, txt(str(value)))

    y -= 19 * mm
    y = section_title(y, "04", "Conformidade e pontos de controlo", "Validacao antes da expedicao")
    quality_text = str(technical.get("controlo_qualidade", "") or "Criterios especificos por definir.")
    standards_text = str(technical.get("normas_conformidade", "") or "Normas aplicaveis por definir.")
    gates = [
        ("01", "INSPECAO DIMENSIONAL", "Conferir dimensoes criticas e tolerancias do desenho."),
        ("02", "MATERIAIS E ACABAMENTO", "Validar materiais, tratamentos, cor e integridade superficial."),
        ("03", "MONTAGEM E INTERFACES", "Confirmar fixacoes, ligacoes e configuracao final."),
        ("04", "ENSAIO FINAL", quality_text),
    ]
    gate_w = (inner_w - 3 * mm) / 2
    for index, (gate_no, gate_name, gate_text) in enumerate(gates):
        col = index % 2
        row = index // 2
        x = margin + col * (gate_w + 3 * mm)
        y_top = y - row * 16 * mm
        canvas_obj.setFillColor(surface)
        canvas_obj.setStrokeColor(line)
        canvas_obj.roundRect(x, y_top - 13 * mm, gate_w, 13 * mm, 2, stroke=1, fill=1)
        canvas_obj.setFillColor(cyan)
        canvas_obj.circle(x + 6 * mm, y_top - 6.5 * mm, 3.2 * mm, stroke=0, fill=1)
        canvas_obj.setFillColor(white)
        canvas_obj.setFont("Helvetica-Bold", 5.8)
        canvas_obj.drawCentredString(x + 6 * mm, y_top - 8.3 * mm, txt(gate_no))
        canvas_obj.setFillColor(ink)
        canvas_obj.setFont("Helvetica-Bold", 6)
        canvas_obj.drawString(x + 12 * mm, y_top - 4.5 * mm, txt(gate_name))
        canvas_obj.setFont("Helvetica", 4.9)
        gate_lines = _pdf_wrap_text(gate_text, "Helvetica", 4.9, gate_w - 15 * mm, max_lines=2) or ["-"]
        gate_line_y = y_top - 8.5 * mm
        for gate_line in gate_lines:
            canvas_obj.drawString(x + 12 * mm, gate_line_y, txt(gate_line))
            gate_line_y -= 2.5 * mm
    canvas_obj.setFillColor(muted)
    canvas_obj.setFont("Helvetica", 5.4)
    canvas_obj.drawString(margin, margin + 9.2 * mm, txt(_pdf_clip_text(f"REFERENCIAL: {standards_text}", inner_w, "Helvetica", 5.4)))
    controlled_footer(1, "CAPA TECNICA")

    columns = [
        ("POS", 10 * mm), ("TIPO", 13 * mm), ("CODIGO / REF.", 27 * mm), ("DESCRICAO", 48 * mm),
        ("MATERIAL / ESPECIFICACAO", 38 * mm), ("QTD", 14 * mm), ("UN.", 10 * mm), ("PROCESSO / DESTINO", inner_w - 160 * mm),
    ]

    def item_kind(item: dict[str, Any]) -> str:
        if ports.orc_line_is_product(item):
            return "STOCK"
        if ports.orc_line_is_service(item):
            return "OPER."
        if str(item.get("stock_material_id", "") or "").strip():
            return "M.PRIMA"
        return "FAB."

    def technical_text(value: object) -> str:
        parts = []
        for part in str(value or "").split("|"):
            clean = part.strip()
            normalized = clean.lower()
            if not clean:
                continue
            if (
                "eur" in normalized
                or "€" in clean
                or "preco" in normalized
                or "preço" in normalized
                or "custo" in normalized
            ):
                continue
            parts.append(clean)
        return " | ".join(parts)

    for bom_page in range(bom_pages):
        canvas_obj.showPage()
        page_number = bom_page + 2
        top = page_h - margin
        canvas_obj.setFillColor(white)
        canvas_obj.setStrokeColor(line)
        canvas_obj.roundRect(margin, top - 27 * mm, inner_w, 27 * mm, 4, stroke=1, fill=1)
        canvas_obj.setFillColor(cyan)
        canvas_obj.rect(margin, top - 27 * mm, 4 * mm, 27 * mm, stroke=0, fill=1)
        canvas_obj.setFillColor(navy)
        canvas_obj.setFont("Helvetica-Bold", 14)
        canvas_obj.drawString(margin + 9 * mm, top - 9 * mm, txt("BOM | BILL OF MATERIALS"))
        header_left_w = inner_w - 53 * mm
        canvas_obj.setFont("Helvetica", 6.4)
        header_lines = _pdf_wrap_text(
            f"{name} | PRODUTO {code} | PARAM {param_code} | {revision}",
            "Helvetica",
            6.4,
            header_left_w,
            max_lines=2,
        )
        header_line_y = top - 15.5 * mm
        for header_line in header_lines:
            canvas_obj.drawString(margin + 9 * mm, header_line_y, txt(header_line))
            header_line_y -= 3.2 * mm
        canvas_obj.setFillColor(green)
        canvas_obj.roundRect(page_w - margin - 45 * mm, top - 12.5 * mm, 38 * mm, 7.5 * mm, 2, stroke=0, fill=1)
        canvas_obj.setFillColor(navy)
        canvas_obj.setFont("Helvetica-Bold", 7)
        canvas_obj.drawCentredString(page_w - margin - 26 * mm, top - 9.8 * mm, txt("MATERIAL + QTD."))
        canvas_obj.setFillColor(muted)
        canvas_obj.setFont("Helvetica", 6)
        canvas_obj.drawRightString(page_w - margin - 7 * mm, top - 16 * mm, txt(f"FOLHA {bom_page + 1}/{bom_pages}"))

        table_top = top - 33 * mm
        canvas_obj.setFillColor(surface)
        canvas_obj.setStrokeColor(line)
        canvas_obj.rect(margin, table_top - 8 * mm, inner_w, 8 * mm, stroke=1, fill=1)
        canvas_obj.setFillColor(navy)
        canvas_obj.setFont("Helvetica-Bold", 5.7)
        x = margin
        for label, width in columns:
            header_label = _pdf_clip_text(label, width - 4, "Helvetica-Bold", 5.7)
            canvas_obj.drawString(x + 2, table_top - 5.2 * mm, txt(header_label))
            x += width
        table_top -= 8 * mm

        page_items = items[bom_page * rows_per_page:(bom_page + 1) * rows_per_page]
        for local_index, item in enumerate(page_items):
            global_index = bom_page * rows_per_page + local_index + 1
            row_h = 10.4 * mm
            row_y = table_top - row_h
            canvas_obj.setFillColor(white if local_index % 2 == 0 else surface)
            canvas_obj.setStrokeColor(line)
            canvas_obj.rect(margin, row_y, inner_w, row_h, stroke=1, fill=1)
            ref = str(item.get("produto_codigo", "") or item.get("stock_material_id", "") or item.get("ref_externa", "") or "-").strip() or "-"
            material = str(item.get("material", "") or "").strip()
            thickness = str(item.get("espessura", "") or "").strip()
            dimensions = technical_text(item.get("dimensao", item.get("dimensoes", "")))
            material_parts = [part for part in (material, f"e={thickness}" if thickness and thickness != "-" else "", dimensions) if part]
            material_spec = " | ".join(material_parts) or "-"
            process = str(item.get("operacao", "") or ("Montagem" if ports.orc_line_is_product(item) else "-")).strip() or "-"
            unit = str(item.get("produto_unid", "") or ("SV" if ports.orc_line_is_service(item) else "UN")).strip() or "UN"
            values = [
                f"{global_index:03d}", item_kind(item), ref, technical_text(item.get("descricao", "")) or "-",
                material_spec, ports._fmt(item.get("qtd", 0)), unit, process,
            ]
            x = margin
            canvas_obj.setFillColor(ink)
            for col_index, (value, (_label, width)) in enumerate(zip(values, columns)):
                font_name = "Helvetica-Bold" if col_index in {0, 1, 2, 5} else "Helvetica"
                font_size = 5.7 if col_index not in {3, 4, 7} else 5.35
                canvas_obj.setFont(font_name, font_size)
                if col_index in {3, 4, 7}:
                    cell_lines = _pdf_wrap_text(value, font_name, font_size, width - 4, max_lines=2) or ["-"]
                else:
                    cell_lines = [_pdf_clip_text(value, width - 4, font_name, font_size)]
                if col_index in {0, 1, 5, 6}:
                    first_line = txt(cell_lines[0])
                    if col_index == 5:
                        canvas_obj.drawRightString(x + width - 2, row_y + 5.8 * mm, first_line)
                    else:
                        canvas_obj.drawCentredString(x + width / 2, row_y + 5.8 * mm, first_line)
                else:
                    cell_y = row_y + 6.5 * mm
                    for cell_line in cell_lines:
                        canvas_obj.drawString(x + 2, cell_y, txt(cell_line))
                        cell_y -= 3.0 * mm
                if col_index < len(columns) - 1:
                    canvas_obj.setStrokeColor(colors.HexColor("#DFE5EA"))
                    canvas_obj.line(x + width, row_y, x + width, row_y + row_h)
                x += width
            table_top = row_y

        summary_y = margin + 11 * mm
        canvas_obj.setFillColor(surface)
        canvas_obj.setStrokeColor(line)
        canvas_obj.roundRect(margin, summary_y, inner_w, 12 * mm, 2.5, stroke=1, fill=1)
        canvas_obj.setFillColor(steel)
        canvas_obj.setFont("Helvetica-Bold", 6.2)
        canvas_obj.drawString(margin + 3 * mm, summary_y + 7.2 * mm, txt("ESTRUTURA DO CONJUNTO"))
        canvas_obj.setFillColor(ink)
        canvas_obj.setFont("Helvetica", 6)
        canvas_obj.drawString(margin + 3 * mm, summary_y + 3.2 * mm, txt(f"{len(items)} referencias | {fabricated_count} fabricadas | {stock_count} stock/materia-prima | {service_count} operacoes"))
        canvas_obj.setFillColor(green)
        canvas_obj.roundRect(page_w - margin - 42 * mm, summary_y + 2 * mm, 39 * mm, 8 * mm, 2, stroke=0, fill=1)
        canvas_obj.setFillColor(navy)
        canvas_obj.setFont("Helvetica-Bold", 6)
        canvas_obj.drawCentredString(page_w - margin - 22.5 * mm, summary_y + 4.8 * mm, txt("BOM CONTROLADA"))
        controlled_footer(page_number, "LISTA DE MATERIAIS")

    canvas_obj.save()
    return target
