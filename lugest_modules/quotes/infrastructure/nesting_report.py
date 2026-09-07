from __future__ import annotations


from datetime import datetime
from pathlib import Path
from typing import Any

from lugest_infra.pdf.text import clip_text as _pdf_clip_text
from lugest_infra.pdf.text import wrap_text as _pdf_wrap_text


from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class NestingReportPorts:
    _draw_operator_logo_plate: Callable[..., Any]
    _fmt: Callable[..., Any]
    _fmt_eur: Callable[..., Any]
    _operator_label_palette: Callable[..., Any]
    _parse_float: Callable[..., Any]
    branding_settings: Callable[..., Any]
    orc_detail: Callable[..., Any]
    orc_nesting_studies: Callable[..., Any]

def render_nesting_study(ports: NestingReportPorts, numero: str, path: str | Path, group_key: str = "") -> Path:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas

    numero_txt = str(numero or "").strip()
    detail = ports.orc_detail(numero_txt)
    studies = ports.orc_nesting_studies(numero_txt)
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
    placements_by_ref: dict[str, list[dict[str, Any]]] = {}
    for sheet in sheets:
        for placement in list(sheet.get("placements", []) or []):
            row = dict(placement or {})
            ref_key = str(row.get("ref_externa", "") or row.get("description", "") or "Sem referência").strip() or "Sem referência"
            placements_by_ref.setdefault(ref_key, []).append(row)
    part_catalog: list[dict[str, Any]] = []
    known_refs: set[str] = set()
    for row in part_rows:
        ref_key = str(row.get("ref_externa", "") or row.get("description", "") or "Sem referência").strip() or "Sem referência"
        matching = placements_by_ref.get(ref_key, [])
        enriched = dict(row)
        enriched["_placements"] = matching
        enriched["_placed_qty"] = len(matching)
        part_catalog.append(enriched)
        known_refs.add(ref_key)
    for ref_key, matching in placements_by_ref.items():
        if ref_key in known_refs:
            continue
        first = dict(matching[0] or {}) if matching else {}
        part_catalog.append(
            {
                "ref_externa": ref_key,
                "description": str(first.get("description", "") or ref_key),
                "qty": len(matching),
                "_placed_qty": len(matching),
                "_placements": matching,
            }
        )

    target = Path(path)
    palette = ports._operator_label_palette()
    branding = ports.branding_settings()
    nesting_logo_txt = str(branding.get("logo_path", "") or "").strip()
    logo_path = Path(nesting_logo_txt) if nesting_logo_txt and Path(nesting_logo_txt).exists() else None
    page_width, page_height = A4
    margin = 26
    report_green = colors.HexColor("#7ED321")
    report_charcoal = colors.HexColor("#3B3B3B")
    report_line = colors.HexColor("#D7D7D7")
    part_cards_per_page = 2
    total_pages = 1 + len(sheets) + max(1, (len(part_catalog) + part_cards_per_page - 1) // part_cards_per_page) + 1
    c = pdf_canvas.Canvas(str(target), pagesize=A4)
    c.setTitle(f"LuGEST - Relatório de Nesting - {numero_txt}")
    c.setAuthor("LuGEST Software ERP Industrial")
    c.setSubject(f"Estudo de nesting, consumo de chapa e custo do orçamento {numero_txt}")
    c.setKeywords("LuGEST, nesting, corte laser, chapa, orçamento, produção")
    page_number = 1

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
        c.setFillColor(colors.white)
        c.rect(0, 0, page_width, page_height, stroke=0, fill=1)
        generated = datetime.now()
        weekdays_pt = [
            "Segunda-feira",
            "Terça-feira",
            "Quarta-feira",
            "Quinta-feira",
            "Sexta-feira",
            "Sábado",
            "Domingo",
        ]
        months_pt = [
            "",
            "janeiro",
            "fevereiro",
            "março",
            "abril",
            "maio",
            "junho",
            "julho",
            "agosto",
            "setembro",
            "outubro",
            "novembro",
            "dezembro",
        ]
        c.setFillColor(report_charcoal)
        set_font(False, 8.5)
        c.drawString(
            margin,
            page_height - 28,
            f"{weekdays_pt[generated.weekday()]}, {generated.day} de {months_pt[generated.month]} de {generated.year}",
        )
        c.setFillColor(report_green)
        set_font(True, 8.5)
        c.drawString(margin, page_height - 42, "Gerado pelo LuGEST Software ERP Industrial")
        ports._draw_operator_logo_plate(
            c,
            palette,
            logo_path,
            page_width - margin - 104,
            page_height - 50,
            104,
            32,
            radius=0,
            padding_x=2,
            padding_y=1,
            line_width=0,
        )
        c.setFillColor(report_charcoal)
        if title == "Relatório de Nesting":
            set_font(True, 21)
            c.drawCentredString(page_width / 2, page_height - 92, f"Estudo de Nesting #{numero_txt}")
            set_font(False, 9)
            c.setFillColor(palette["muted"])
            c.drawCentredString(page_width / 2, page_height - 108, _pdf_clip_text(subtitle, page_width - 90, font_regular, 9))
            return page_height - 136
        set_font(True, 19)
        c.drawString(margin, page_height - 82, title)
        set_font(False, 8.5)
        c.setFillColor(palette["muted"])
        c.drawString(margin, page_height - 99, _pdf_clip_text(subtitle, page_width - (margin * 2), font_regular, 8.5))
        return page_height - 120

    def draw_footer() -> None:
        c.setStrokeColor(report_line)
        c.setLineWidth(0.7)
        c.line(0, 34, page_width, 34)
        c.setFillColor(report_charcoal)
        set_font(True, 8)
        c.drawString(margin, 19, f"Estudo #{numero_txt} - {datetime.now().strftime('%d/%m/%Y')}")
        set_font(False, 8)
        c.drawCentredString(page_width / 2, 19, f"Página {page_number} / {total_pages}")
        c.setFillColor(report_green)
        set_font(True, 8)
        c.drawRightString(page_width - margin, 19, "LuGEST")

    def start_page(title: str, subtitle: str) -> float:
        nonlocal page_number
        draw_footer()
        c.showPage()
        page_number += 1
        return draw_header(title, subtitle)

    def draw_metric_card(x: float, y_top: float, width: float, title: str, value: str, accent: Any) -> None:
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(x, y_top - 52, width, 46, 3, stroke=1, fill=1)
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
        c.roundRect(x, y_top - box_height, width, box_height, 3, stroke=1, fill=1)
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
        return start_page(title, subtitle)

    def draw_table_header(y_top: float, columns: list[tuple[str, float]]) -> tuple[float, list[float], float]:
        total_width = page_width - (margin * 2)
        c.setFillColor(palette["surface_alt"])
        c.setStrokeColor(palette["line"])
        c.rect(margin, y_top - 20, total_width, 18, stroke=1, fill=1)
        c.setFillColor(palette["primary_dark"])
        set_font(True, 8)
        x_positions: list[float] = []
        cursor_x = margin + 7
        for label, ratio in columns:
            x_positions.append(cursor_x)
            c.drawString(cursor_x, y_top - 13, label)
            cursor_x += total_width * ratio
        return y_top - 24, x_positions, total_width

    def draw_key_value_table(
        y_top: float,
        title: str,
        rows: list[tuple[str, str]],
        *,
        value_accent_rows: set[int] | None = None,
    ) -> float:
        value_accent_rows = set(value_accent_rows or set())
        c.setFillColor(report_charcoal)
        set_font(False, 8.2)
        c.drawCentredString(page_width / 2, y_top, title.upper())
        y_cursor = y_top - 18
        label_width = (page_width - (margin * 2)) * 0.60
        for row_index, (label, value) in enumerate(rows):
            c.setStrokeColor(report_line)
            c.setLineWidth(0.6)
            c.line(margin, y_cursor - 7, page_width - margin, y_cursor - 7)
            c.setFillColor(report_charcoal)
            set_font(True, 9)
            c.drawString(margin + 6, y_cursor, label)
            c.setFillColor(report_green if row_index in value_accent_rows else report_charcoal)
            set_font(True if row_index in value_accent_rows else False, 9)
            c.drawString(margin + label_width, y_cursor, value)
            y_cursor -= 25
        return y_cursor - 12

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
            f"{ports._fmt(float(sheet.get('sheet_width_mm', 0) or 0))} x "
            f"{ports._fmt(float(sheet.get('sheet_height_mm', 0) or 0))} mm | "
            f"origem {str(sheet.get('source_kind', '-') or '-').strip()} | "
            f"real {ports._fmt(sheet.get('utilization_net_pct', 0))}%"
        )
        c.drawString(
            x + 10,
            y_top - 28,
            _pdf_clip_text(tech_line, width - 20, font_regular, 7.8),
        )
        c.drawString(
            x + 10,
            y_top - 40,
            f"{int(sheet.get('part_count', 0) or 0)} peca(s) | bbox {ports._fmt(sheet.get('utilization_bbox_pct', 0))}% | compra {'sim' if str(sheet.get('source_kind', '') or '').strip().lower() == 'purchase' else 'nao'}",
        )
        legend_rows = [dict(row or {}) for row in list(sheet.get("placements", []) or [])]
        legend_groups: list[dict[str, Any]] = []
        legend_by_ref: dict[str, dict[str, Any]] = {}
        for placement in legend_rows:
            ref_txt = str(placement.get("ref_externa", "") or placement.get("description", "") or "-").strip() or "-"
            group = legend_by_ref.get(ref_txt)
            if group is None:
                group = {"id": len(legend_groups) + 1, "ref": ref_txt, "qty": 0}
                legend_by_ref[ref_txt] = group
                legend_groups.append(group)
            group["qty"] = int(group.get("qty", 0) or 0) + 1
        inner_x = x + 12
        inner_h = height - 66
        panel_gap = 10.0
        if page_height > page_width:
            map_x = inner_x
            map_w = width - 24
            map_h = min(320.0, max(220.0, inner_h * 0.55))
            map_y = y_top - 52 - map_h
            legend_x = inner_x
            legend_w = width - 24
            legend_h = 112.0
            legend_y_bottom = map_y - panel_gap - legend_h
        else:
            legend_x = inner_x
            legend_y_bottom = y_top - height + 14
            legend_w = max(154.0, min(210.0, width * 0.25))
            legend_h = max(80.0, inner_h)
            map_x = inner_x + legend_w + panel_gap
            map_w = width - 24 - legend_w - panel_gap
            map_y = legend_y_bottom
            map_h = max(80.0, inner_h)
        c.setFillColor(colors.HexColor("#F9FBFE"))
        c.roundRect(legend_x, legend_y_bottom, legend_w, legend_h, 12, stroke=0, fill=1)
        c.setStrokeColor(palette["line"])
        c.roundRect(legend_x, legend_y_bottom, legend_w, legend_h, 12, stroke=1, fill=0)
        c.setFillColor(colors.HexColor("#F8FAFC"))
        c.roundRect(map_x, map_y, map_w, map_h, 12, stroke=0, fill=1)
        c.setStrokeColor(colors.HexColor("#9cb2c8"))
        c.roundRect(map_x, map_y, map_w, map_h, 12, stroke=1, fill=0)

        set_font(True, 8.4)
        c.setFillColor(palette["ink"])
        legend_title_y = legend_y_bottom + legend_h - 14
        c.drawString(legend_x + 10, legend_title_y, "Peças no layout")
        set_font(False, 7.2)
        legend_y = legend_title_y - 14
        max_legend_rows = min(len(legend_groups), 16)
        for legend_index, group in enumerate(legend_groups[:max_legend_rows]):
            prefix = f"{int(group.get('id', 0) or 0):02d}"
            ref_txt = str(group.get("ref", "") or "-")
            qty_txt = int(group.get("qty", 0) or 0)
            if page_height > page_width:
                legend_column = legend_index // 8
                legend_row = legend_index % 8
                legend_draw_x = legend_x + 10 + (legend_column * (legend_w / 2.0))
                legend_draw_y = legend_title_y - 14 - (legend_row * 10)
                legend_text_w = (legend_w / 2.0) - 42
            else:
                legend_draw_x = legend_x + 10
                legend_draw_y = legend_y
                legend_text_w = legend_w - 40
            c.setFillColor(colors.HexColor("#0f172a"))
            c.drawString(legend_draw_x, legend_draw_y, prefix)
            c.drawString(
                legend_draw_x + 18,
                legend_draw_y,
                _pdf_clip_text(f"{ref_txt}  x{qty_txt}", legend_text_w, font_regular, 7.2),
            )
            if page_height <= page_width:
                legend_y -= 10
        remaining = len(legend_groups) - max_legend_rows
        if remaining > 0:
            c.setFillColor(palette["muted"])
            c.drawString(legend_x + 10, legend_y, f"+{remaining} referência(s) adicionais")

        if page_height > page_width:
            info_y_bottom = y_top - height + 14
            info_y_top = legend_y_bottom - panel_gap
            info_h = max(90.0, info_y_top - info_y_bottom)
            c.setFillColor(colors.white)
            c.setStrokeColor(palette["line"])
            c.roundRect(inner_x, info_y_bottom, width - 24, info_h, 12, stroke=1, fill=1)
            c.setFillColor(palette["ink"])
            set_font(True, 8.4)
            c.drawString(inner_x + 10, info_y_top - 16, "Informação do layout")
            source_kind = str(sheet.get("source_kind", "") or "").strip().lower()
            source_label = {"purchase": "Compra", "stock": "Stock", "remnant": "Retalho"}.get(source_kind, source_kind.title() or "-")
            info_rows = [
                ("Dimensões da chapa", f"{ports._fmt(sheet.get('sheet_width_mm', 0))} x {ports._fmt(sheet.get('sheet_height_mm', 0))} mm"),
                ("Peças colocadas", str(int(sheet.get("part_count", 0) or 0))),
                ("Utilização real", f"{ports._fmt(sheet.get('utilization_net_pct', 0))}%"),
                ("Ocupação delimitadora", f"{ports._fmt(sheet.get('utilization_bbox_pct', 0))}%"),
                ("Origem", source_label),
                ("Necessita compra", "Sim" if source_kind == "purchase" else "Não"),
            ]
            row_y = info_y_top - 34
            row_w = (width - 44) / 2.0
            for info_index, (label, value) in enumerate(info_rows):
                info_column = info_index % 2
                info_row = info_index // 2
                draw_x = inner_x + 10 + (info_column * row_w)
                draw_y = row_y - (info_row * 24)
                c.setFillColor(palette["muted"])
                set_font(True, 6.8)
                c.drawString(draw_x, draw_y, label.upper())
                c.setFillColor(palette["ink"])
                set_font(False, 8)
                c.drawString(draw_x, draw_y - 10, _pdf_clip_text(value, row_w - 12, font_regular, 8))

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
        c.setStrokeColor(colors.HexColor("#A0A0A0"))
        c.setFillColor(colors.HexColor("#B7B7B7"))
        if outer_polygons:
            for polygon_points in outer_polygons:
                mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                if len(mapped) >= 3:
                    draw_polygon(mapped, stroke_color=colors.HexColor("#A0A0A0"), fill_color=colors.HexColor("#B7B7B7"), stroke_width=1.1)
            for polygon_points in hole_polygons:
                mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                if len(mapped) >= 3:
                    draw_polygon(mapped, stroke_color=palette["line_strong"], fill_color=colors.white, stroke_width=0.9)
        else:
            c.rect(offset_x, offset_y, body_w, body_h, stroke=1, fill=1)

        palette_hexes = ["#AD659C"]
        for placement_index, placement in enumerate(list(sheet.get("placements", []) or []), start=1):
            ref_txt = str(placement.get("ref_externa", "") or placement.get("description", "") or "-").strip() or "-"
            type_index = int(dict(legend_by_ref.get(ref_txt, {}) or {}).get("id", placement_index) or placement_index)
            c.setStrokeColor(colors.white)
            c.setFillColor(colors.HexColor(palette_hexes[type_index % len(palette_hexes)]))
            poly_groups = [list(points or []) for points in list(placement.get("shape_outer_polygons", []) or [])]
            label_x = offset_x + (float(placement.get("x_mm", 0) or 0) * scale)
            label_y = offset_y + (float(placement.get("y_mm", 0) or 0) * scale)
            label_w = max(1.0, float(placement.get("width_mm", 0) or 0) * scale)
            label_h = max(1.0, float(placement.get("height_mm", 0) or 0) * scale)
            if poly_groups:
                for polygon_points in poly_groups:
                    mapped = [map_point(float(point[0]), float(point[1])) for point in list(polygon_points or []) if isinstance(point, (list, tuple)) and len(point) >= 2]
                    if len(mapped) >= 3:
                        draw_polygon(mapped, stroke_color=colors.white, fill_color=colors.HexColor(palette_hexes[type_index % len(palette_hexes)]), stroke_width=1.2)
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
                c.drawCentredString(label_x + (label_w / 2.0), label_y + (label_h / 2.0), str(type_index))

    def draw_part_thumbnail(x: float, y: float, width: float, height: float, row: dict[str, Any]) -> None:
        placements = [dict(item or {}) for item in list(row.get("_placements", []) or [])]
        placement = placements[0] if placements else {}
        outer = [list(points or []) for points in list(placement.get("shape_outer_polygons", []) or [])]
        holes = [list(points or []) for points in list(placement.get("shape_hole_polygons", []) or [])]
        c.setFillColor(colors.HexColor("#F3F6F8"))
        c.setStrokeColor(palette["line"])
        c.roundRect(x, y, width, height, 8, stroke=1, fill=1)

        valid_points = [
            (float(point[0]), float(point[1]))
            for polygon in outer
            for point in polygon
            if isinstance(point, (list, tuple)) and len(point) >= 2
        ]
        if not valid_points:
            piece_w = max(1.0, float(placement.get("width_mm", row.get("width_mm", 1)) or 1))
            piece_h = max(1.0, float(placement.get("height_mm", row.get("height_mm", 1)) or 1))
            valid_points = [(0.0, 0.0), (piece_w, 0.0), (piece_w, piece_h), (0.0, piece_h)]
            outer = [valid_points]

        min_x = min(point[0] for point in valid_points)
        max_x = max(point[0] for point in valid_points)
        min_y = min(point[1] for point in valid_points)
        max_y = max(point[1] for point in valid_points)
        shape_w = max(1.0, max_x - min_x)
        shape_h = max(1.0, max_y - min_y)
        pad = 10.0
        scale = min((width - pad * 2) / shape_w, (height - pad * 2) / shape_h)
        offset_x = x + ((width - (shape_w * scale)) / 2.0)
        offset_y = y + ((height - (shape_h * scale)) / 2.0)

        def mapped_polygon(points: list[Any]) -> list[tuple[float, float]]:
            return [
                (offset_x + ((float(point[0]) - min_x) * scale), offset_y + ((float(point[1]) - min_y) * scale))
                for point in points
                if isinstance(point, (list, tuple)) and len(point) >= 2
            ]

        def paint_polygon(points: list[tuple[float, float]], *, hole: bool = False) -> None:
            if len(points) < 3:
                return
            path_obj = c.beginPath()
            path_obj.moveTo(points[0][0], points[0][1])
            for px, py in points[1:]:
                path_obj.lineTo(px, py)
            path_obj.close()
            c.setStrokeColor(colors.HexColor("#6F8798"))
            c.setFillColor(colors.white if hole else colors.HexColor("#AEC0CF"))
            c.setLineWidth(0.8)
            c.drawPath(path_obj, stroke=1, fill=1)

        for polygon in outer:
            paint_polygon(mapped_polygon(polygon))
        for polygon in holes:
            paint_polygon(mapped_polygon(polygon), hole=True)

    def draw_part_card(x: float, y_top: float, width: float, height: float, row: dict[str, Any], index: int) -> None:
        c.setFillColor(colors.white)
        c.setStrokeColor(report_line)
        c.roundRect(x, y_top - height, width, height, 4, stroke=1, fill=1)
        ref_txt = str(row.get("ref_externa", "") or row.get("description", "") or "Sem referência").strip() or "Sem referência"
        c.setFillColor(report_charcoal)
        set_font(True, 11)
        c.drawString(x + 12, y_top - 20, f"Peça {index}")
        c.setFillColor(report_green)
        set_font(True, 8)
        c.drawRightString(x + width - 12, y_top - 19, f"ID {index}")

        thumb_w = min(150.0, width * 0.31)
        thumb_h = height - 54
        draw_part_thumbnail(x + 12, y_top - height + 12, thumb_w, thumb_h, row)
        info_x = x + thumb_w + 26
        info_w = width - thumb_w - 38
        value_x = info_x + (info_w * 0.56)
        cursor_y = y_top - 44
        placement = dict(list(row.get("_placements", []) or [{}])[0] or {})
        piece_w = float(placement.get("width_mm", row.get("width_mm", 0)) or 0)
        piece_h = float(placement.get("height_mm", row.get("height_mm", 0)) or 0)
        placed_qty = int(row.get("_placed_qty", 0) or 0)
        requested_qty = int(row.get("qty", placed_qty) or placed_qty)
        details = [
            ("Nome", ref_txt),
            ("Quantidade", f"{placed_qty} / {requested_qty} colocada(s)"),
            ("Dimensões da peça", f"{ports._fmt(piece_w)} x {ports._fmt(piece_h)} mm" if piece_w and piece_h else "-"),
            ("Comprimento de corte", f"{ports._fmt(row.get('cut_length_m', 0))} m"),
            ("Perfurações", str(int(row.get("pierce_count", 0) or 0))),
            ("Tempo de máquina", f"{ports._fmt(row.get('machine_total_min', 0))} min"),
            ("Valor orçamentado", ports._fmt_eur(row.get("quoted_total_eur", 0))),
        ]
        row_h = 24.0
        for row_index, (label, value) in enumerate(details):
            c.setFillColor(colors.white if row_index % 2 == 0 else colors.HexColor("#FAFAFA"))
            c.setStrokeColor(report_line)
            c.rect(info_x, cursor_y - 15, info_w, row_h, stroke=1, fill=1)
            c.setFillColor(report_charcoal)
            set_font(True, 7.6)
            c.drawString(info_x + 7, cursor_y - 1, label)
            set_font(False, 7.6)
            c.drawString(value_x, cursor_y - 1, _pdf_clip_text(value, (info_x + info_w) - value_x - 7, font_regular, 7.6))
            cursor_y -= row_h

    def draw_section_title(y_top: float, title: str, subtitle_text: str = "") -> float:
        c.setFillColor(report_charcoal)
        set_font(True, 14)
        c.drawString(margin, y_top, title)
        if subtitle_text:
            c.setFillColor(palette["muted"])
            set_font(False, 8.5)
            c.drawRightString(page_width - margin, y_top, subtitle_text)
        return y_top - 26

    group_label = str(study.get("group_label", "") or selected_key).strip() or selected_key
    profile_name = str(dict(summary.get("selected_sheet_profile", {}) or {}).get("name", "") or bridge.get("selected_profile_name", "") or "Apenas stock").strip() or "Apenas stock"
    subtitle = f"Orçamento {numero_txt} | Grupo {group_label} | Perfil {profile_name}"
    y = draw_header("Relatório de Nesting", subtitle)

    placed_count = int(bridge.get("part_count_placed", summary.get("part_count_placed", 0)) or 0)
    requested_count = int(bridge.get("part_count_requested", summary.get("part_count_requested", 0)) or 0)
    total_sheet_area_m2 = float(summary.get("total_sheet_area_mm2", 0) or 0) / 1_000_000.0
    parts_area_m2 = float(summary.get("total_part_area_mm2", summary.get("used_net_area_mm2", 0)) or 0) / 1_000_000.0
    thickness_txt = str(bridge.get("espessura", bridge.get("thickness_mm", "")) or "").strip()
    if not thickness_txt and "|" in selected_key:
        thickness_txt = selected_key.rsplit("|", 1)[-1].strip()
    global_rows = [
        ("Eficiência do nesting", f"{ports._fmt(summary.get('utilization_net_pct', 0))}%"),
        ("Número de chapas", str(int(bridge.get("sheet_count", summary.get("sheet_count", len(sheets))) or len(sheets)))),
        ("Número de layouts", str(len(sheets))),
        ("Comprimento de corte / marcação", f"{ports._fmt(totals.get('cut_length_m', 0))} m / {ports._fmt(totals.get('marking_length_m', 0))} m"),
        ("Superfície das chapas / peças", f"{ports._fmt(total_sheet_area_m2)} m² / {ports._fmt(parts_area_m2)} m²"),
        ("Espessura", f"{thickness_txt or '-'} mm"),
        ("Peças colocadas / programadas", f"{placed_count} / {requested_count}"),
    ]
    y = draw_key_value_table(y, "Estatísticas globais", global_rows, value_accent_rows={0, 1})

    optimization_seconds = int(
        ports._parse_float(
            options.get(
                "optimization_time_sec",
                options.get(
                    "time_limit_sec",
                    options.get(
                        "search_seconds",
                        summary.get("optimization_time_limit_s", summary.get("optimization_elapsed_s", 0)),
                    ),
                ),
            ),
            0,
        )
    )
    settings_rows = [
        ("Tempo de otimização", f"{optimization_seconds} segundos"),
        ("Distância entre peças", f"{ports._fmt(options.get('part_spacing_mm', 0))} mm"),
        ("Margem da chapa", f"{ports._fmt(options.get('edge_margin_mm', 0))} mm"),
        ("Espelhamento", "Permitido" if bool(options.get("allow_mirror")) else "Não permitido"),
        ("Rotações", "Automáticas" if bool(options.get("allow_rotate")) else "Bloqueadas"),
    ]
    y = draw_key_value_table(y, "Definições do nesting", settings_rows)

    production_rows = [
        ("Valor comercial das peças", ports._fmt_eur(totals.get("quoted_total_eur", bridge.get("quoted_total_eur", 0)))),
        ("Tempo total de máquina", f"{ports._fmt(totals.get('machine_total_min', 0))} min"),
        ("Perfurações", str(int(totals.get("pierce_count", 0) or 0))),
        ("Chapas de stock / retalhos / compra", f"{int(summary.get('stock_sheet_count', 0) or 0)} / {int(summary.get('remnant_sheet_count', 0) or 0)} / {int(summary.get('purchased_sheet_count', 0) or 0)}"),
    ]
    y = draw_key_value_table(y, "Produção e custo", production_rows, value_accent_rows={0})

    c.setFillColor(report_charcoal)
    set_font(True, 14)
    c.drawString(margin, y, "Conteúdo do relatório")
    navigation = [
        ("1.", "Layouts de nesting"),
        ("2.", "Peças"),
        ("3.", "Chapas utilizadas e decisão"),
    ]
    y -= 27
    for number_txt, label_txt in navigation:
        c.setFillColor(report_charcoal)
        set_font(False, 9)
        c.drawString(margin + 8, y, number_txt)
        c.setFillColor(report_green)
        set_font(True, 9)
        c.drawCentredString(page_width / 2, y, label_txt)
        y -= 17

    if unplaced:
        y = ensure_page(y, 80, "Relatório de Nesting", subtitle)
        unplaced_preview = [
            f"{str(row.get('ref_externa', '-') or '-').strip() or '-'} | {str(row.get('description', '-') or '-').strip() or '-'}"
            for row in unplaced[:6]
        ]
        box_h = draw_info_box(margin, y, page_width - (margin * 2), "Peças fora do plano", unplaced_preview, tone="warning")
        y -= box_h + 10

    for sheet_index, sheet in enumerate(sheets, start=1):
        y = start_page("Layouts de Nesting", subtitle)
        y = draw_section_title(
            y,
            f"Layout {sheet_index} de {len(sheets)}",
            f"Chapa {int(sheet.get('index', sheet_index) or sheet_index)}",
        )
        draw_sheet_map(
            margin,
            y,
            page_width - (margin * 2),
            max(320.0, y - 48),
            sheet,
        )

    cards_per_page = part_cards_per_page if page_height > page_width else 6
    card_gap = 10.0
    catalog_inner_w = page_width - (margin * 2)
    catalog_columns = 1 if page_height > page_width else 2
    catalog_card_w = (catalog_inner_w - (card_gap * (catalog_columns - 1))) / catalog_columns
    catalog_card_h = 252.0 if page_height > page_width else 126.0
    for page_start in range(0, len(part_catalog), cards_per_page):
        y = start_page("Peças do Estudo", subtitle)
        page_rows = part_catalog[page_start : page_start + cards_per_page]
        y = draw_section_title(
            y,
            "Catálogo visual de peças",
            f"{page_start + 1}-{page_start + len(page_rows)} de {len(part_catalog)}",
        )
        for local_index, row in enumerate(page_rows):
            column = local_index % catalog_columns
            row_index = local_index // catalog_columns
            card_x = margin + (column * (catalog_card_w + card_gap))
            card_y_top = y - (row_index * (catalog_card_h + card_gap))
            draw_part_card(
                card_x,
                card_y_top,
                catalog_card_w,
                catalog_card_h,
                row,
                page_start + local_index + 1,
            )

    y = start_page("Chapas e Decisão Final", subtitle)
    y = draw_section_title(
        y,
        "Chapas utilizadas",
        f"{len(sheets)} chapa(s) | {ports._fmt(float(summary.get('total_sheet_area_mm2', 0) or 0) / 1_000_000.0)} m²",
    )
    sheet_columns = [
        ("Chapa", 0.09),
        ("Origem / perfil", 0.31),
        ("Dimensões", 0.23),
        ("Peças", 0.10),
        ("Utilização", 0.16),
        ("Compra", 0.11),
    ]
    y, x_positions, total_w = draw_table_header(y, sheet_columns)
    for row_index, sheet in enumerate(sheets):
        row_h = 22.0
        fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
        c.setFillColor(fill_color)
        c.setStrokeColor(palette["line"])
        c.rect(margin, y - row_h + 2, page_width - (margin * 2), row_h - 2, stroke=1, fill=1)
        source_kind = str(sheet.get("source_kind", "") or "").strip().lower()
        values = [
            str(int(sheet.get("index", row_index + 1) or row_index + 1)),
            str(sheet.get("source_label", "") or profile_name),
            f"{ports._fmt(sheet.get('sheet_width_mm', 0))} x {ports._fmt(sheet.get('sheet_height_mm', 0))} mm",
            str(int(sheet.get("part_count", 0) or 0)),
            f"{ports._fmt(sheet.get('utilization_net_pct', 0))}%",
            "Sim" if source_kind == "purchase" else "Não",
        ]
        c.setFillColor(palette["ink"])
        set_font(False, 7.5)
        for idx, value in enumerate(values):
            max_w = (total_w * sheet_columns[idx][1]) - 10
            c.drawString(x_positions[idx], y - 13, _pdf_clip_text(value, max_w, font_regular, 7.5))
        y -= row_h
    y -= 14

    if sheet_candidates:
        y = draw_section_title(y, "Cenários analisados", f"{len(sheet_candidates)} alternativa(s)")
        candidate_columns = [
            ("Cenário", 0.31),
            ("Método", 0.28),
            ("Chapas", 0.10),
            ("Compactação", 0.15),
            ("Compra", 0.16),
        ]
        y, x_positions, total_w = draw_table_header(y, candidate_columns)
        for row_index, candidate in enumerate(sheet_candidates[:7]):
            row_h = 19.0
            fill_color = colors.white if row_index % 2 == 0 else palette["surface_alt"]
            c.setFillColor(fill_color)
            c.setStrokeColor(palette["line"])
            c.rect(margin, y - row_h + 2, page_width - (margin * 2), row_h - 2, stroke=1, fill=1)
            values = [
                str(candidate.get("name", "") or "-"),
                str(candidate.get("method", "") or bridge.get("analysis_method", "-")),
                str(int(candidate.get("sheet_count", 0) or 0)),
                f"{ports._fmt(candidate.get('layout_compactness_pct', 0))}%",
                f"{ports._fmt((candidate.get('purchase_sheet_area_mm2', 0) or 0) / 1_000_000.0)} m²",
            ]
            c.setFillColor(palette["ink"])
            set_font(False, 7.4)
            for idx, value in enumerate(values):
                max_w = (total_w * candidate_columns[idx][1]) - 10
                c.drawString(x_positions[idx], y - 12, _pdf_clip_text(value, max_w, font_regular, 7.4))
            y -= row_h
        y -= 12

    final_notes = []
    if unplaced:
        final_notes.append(f"Peças não colocadas: {len(unplaced)}")
    final_notes.extend(warnings[:4])
    final_notes.extend(decision_lines[:3])
    if not final_notes:
        final_notes = ["Estudo concluído sem pendências técnicas registadas."]
    draw_info_box(
        margin,
        y,
        page_width - (margin * 2),
        "Conclusão e pendências",
        final_notes,
        tone="warning" if unplaced or warnings else "info",
    )

    draw_footer()
    c.save()
    return target
