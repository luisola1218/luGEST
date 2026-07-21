from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from reportlab.lib import colors

from .text import clip_text, fit_font_size, wrap_text


REGULAR = "Helvetica"
BOLD = "Helvetica-Bold"


def _text(value: Any) -> str:
    return str(value if value not in (None, "") else "-").strip() or "-"


def _page(c, width: float, height: float) -> None:
    c.setFillColor(colors.white)
    c.rect(0, 0, width, height, stroke=0, fill=1)


def _box(c, x: float, y: float, width: float, height: float, palette, *, fill=None, radius: float = 4) -> None:
    c.setFillColor(fill or palette["surface"])
    c.setStrokeColor(palette["line"])
    c.setLineWidth(0.75)
    c.roundRect(x, y, width, height, radius, stroke=1, fill=1)


def _label_value(
    c,
    x: float,
    y: float,
    width: float,
    height: float,
    label: str,
    value: str,
    palette,
    *,
    value_size: float = 8,
    fill=None,
    max_value_size: float | None = None,
) -> None:
    _box(c, x, y, width, height, palette, fill=fill)
    compact = height <= 22
    label_size = 4.4 if compact else 5.6
    label_y = y + height - (6.5 if compact else 10)
    value_y = y + (3 if compact else 6)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, label_size)
    c.drawString(x + 7, label_y, clip_text(label, width - 14, REGULAR, label_size))
    preferred = min(float(value_size), float(max_value_size) if max_value_size is not None else (6.8 if compact else 8.0))
    size = fit_font_size(value, BOLD, width - 14, preferred, 5.1 if compact else 6)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, size)
    c.drawString(x + 7, value_y, clip_text(value, width - 14, BOLD, size))


def draw_product_label(
    c,
    width: float,
    height: float,
    product: dict[str, Any],
    palette: dict[str, Any],
    logo_path: Path | None,
    printed_at: str,
    *,
    draw_logo: Callable[..., None],
    draw_barcode: Callable[..., None],
    scan_code: str,
    quantity_text: str,
    price_text: str,
    dimension_text: str,
) -> None:
    _page(c, width, height)
    outer = 6
    c.setFillColor(colors.white)
    c.setStrokeColor(palette["line_strong"])
    c.setLineWidth(1)
    c.roundRect(outer, outer, width - 2 * outer, height - 2 * outer, 5, stroke=1, fill=1)
    c.setFillColor(palette["primary"])
    c.rect(outer, height - outer - 2.5, width - 2 * outer, 2.5, stroke=0, fill=1)

    x, body_w = outer + 5, width - 2 * outer - 10
    top = height - outer - 5
    header_h = 24
    header_y = top - header_h
    draw_logo(c, palette, logo_path, x, header_y + 3, 44, 18, radius=4, padding_x=2, padding_y=1.5, line_width=0.65)
    stock_w = 66
    stock_x = x + body_w - stock_w
    _label_value(c, stock_x, header_y + 2, stock_w, 20, "Stock", quantity_text, palette, value_size=9.2, max_value_size=9.2, fill=palette["primary_soft"])
    title_x = x + 50
    title_w = stock_x - title_x - 6
    title = "Etiqueta de Produto"
    title_size = fit_font_size(title, BOLD, title_w, 9.5, 7.2)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, title_size)
    c.drawString(title_x, header_y + 13.5, clip_text(title, title_w, BOLD, title_size))
    c.setFillColor(palette["muted"])
    c.setFont(BOLD, 6.8)
    c.drawString(title_x, header_y + 3.5, clip_text(_text(product.get("codigo")), title_w, BOLD, 6.8))

    identity_h = 33
    identity_y = header_y - identity_h - 5
    _box(c, x, identity_y, body_w, identity_h, palette, fill=colors.white, radius=4)
    code = _text(product.get("codigo"))
    code_size = fit_font_size(code, BOLD, body_w - 14, 13, 9)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, code_size)
    c.drawString(x + 7, identity_y + 17, clip_text(code, body_w - 14, BOLD, code_size))
    description = _text(product.get("descricao"))
    desc_size = fit_font_size(description, REGULAR, body_w - 14, 8.2, 6.4)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, desc_size)
    c.drawString(x + 7, identity_y + 5.5, clip_text(description, body_w - 14, REGULAR, desc_size))

    info_h, gap = 19, 5
    info_y = identity_y - info_h - 5
    info_w = (body_w - gap) / 2
    meta = " | ".join((_text(product.get("categoria")), _text(product.get("tipo")), _text(dimension_text)))
    _label_value(c, x, info_y, info_w, info_h, "Categoria / tipo / dimensao", meta, palette, value_size=7.6, max_value_size=7.6, fill=palette["surface_alt"])
    _label_value(c, x + info_w + gap, info_y, info_w, info_h, "Preco / atualizado", f"{price_text} | {printed_at[:10]}", palette, value_size=7.6, max_value_size=7.6)

    barcode_y = outer + 4
    barcode_h = info_y - barcode_y - 5
    _box(c, x, barcode_y, body_w, barcode_h, palette, fill=colors.white, radius=4)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 5.8)
    c.drawString(x + 6, barcode_y + barcode_h - 7, "Codigo para picagem e selecao")
    bar_x, bar_w = x + 6, body_w - 12
    bar_h = max(12, barcode_h - 19)
    draw_barcode(c, scan_code, bar_x, barcode_y + 7, bar_w, bar_h, min_bar_width=0.5, max_bar_width=1.6)
    code_size = fit_font_size(scan_code, BOLD, bar_w, 7.5, 6.2)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, code_size)
    c.drawCentredString(x + body_w / 2, barcode_y + 1.5, clip_text(scan_code, bar_w, BOLD, code_size))


def draw_material_label(
    c,
    width: float,
    height: float,
    record: dict[str, Any],
    palette: dict[str, Any],
    logo_path: Path | None,
    printed_at: str,
    *,
    draw_logo: Callable[..., None],
    draw_barcode: Callable[..., None],
    scan_code: str,
    dimension_text: str,
    lot_text: str,
    location_text: str,
    available_text: str,
    format_text: str,
    kind_text: str,
    weight_text: str,
) -> None:
    _page(c, width, height)
    outer = 6
    c.setFillColor(colors.white)
    c.setStrokeColor(palette["line_strong"])
    c.setLineWidth(1)
    c.roundRect(outer, outer, width - 2 * outer, height - 2 * outer, 5, stroke=1, fill=1)
    c.setFillColor(palette["primary"])
    c.rect(outer, height - outer - 2.5, width - 2 * outer, 2.5, stroke=0, fill=1)
    x, body_w = outer + 5, width - 2 * outer - 10
    top = height - outer - 5

    header_h = 24
    header_y = top - header_h
    draw_logo(c, palette, logo_path, x, header_y + 3, 44, 18, radius=4, padding_x=2, padding_y=1.5, line_width=0.65)
    stock_w = 66
    stock_x = x + body_w - stock_w
    _label_value(c, stock_x, header_y + 2, stock_w, 20, "Stock disponivel", available_text, palette, value_size=9.2, max_value_size=9.2, fill=palette["primary_soft"])
    title_x = x + 50
    title_w = stock_x - title_x - 6
    title = "Etiqueta de Materia-Prima"
    title_size = fit_font_size(title, BOLD, title_w, 9.5, 7.2)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, title_size)
    c.drawString(title_x, header_y + 13.5, clip_text(title, title_w, BOLD, title_size))
    c.setFillColor(palette["muted"])
    c.setFont(BOLD, 6.8)
    c.drawString(title_x, header_y + 3.5, clip_text(_text(record.get("id")), title_w, BOLD, 6.8))

    hero_h = 33
    hero_y = header_y - hero_h - 5
    _box(c, x, hero_y, body_w, hero_h, palette, fill=colors.white, radius=4)
    material = _text(record.get("material"))
    thickness = _text(record.get("espessura"))
    hero = f"{material} | {thickness} mm"
    hero_size = fit_font_size(hero, BOLD, body_w - 14, 13.5, 9)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, hero_size)
    c.drawString(x + 7, hero_y + 17, clip_text(hero, body_w - 14, BOLD, hero_size))
    identity = f"{kind_text} | {format_text} | {weight_text}"
    identity_size = fit_font_size(identity, REGULAR, body_w - 14, 8, 6.4)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, identity_size)
    c.drawString(x + 7, hero_y + 5.5, clip_text(identity, body_w - 14, REGULAR, identity_size))

    meta_h, gap = 21, 5
    meta_y = hero_y - meta_h - 5
    dim_w = body_w * 0.50
    lot_w = body_w * 0.25
    location_w = body_w - dim_w - lot_w - 2 * gap
    _label_value(c, x, meta_y, dim_w, meta_h, "Dimensao identificada", dimension_text, palette, value_size=8.4, max_value_size=8.4, fill=palette["surface_alt"])
    _label_value(c, x + dim_w + gap, meta_y, lot_w, meta_h, "Lote", lot_text, palette, value_size=7.4, max_value_size=7.4)
    _label_value(c, x + dim_w + lot_w + 2 * gap, meta_y, location_w, meta_h, "Localizacao", location_text, palette, value_size=7.4, max_value_size=7.4)

    barcode_y = outer + 4
    barcode_h = meta_y - barcode_y - 5
    _box(c, x, barcode_y, body_w, barcode_h, palette, fill=colors.white, radius=4)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 5.8)
    c.drawString(x + 6, barcode_y + barcode_h - 7, "Codigo para picagem e identificacao")
    c.drawRightString(x + body_w - 6, barcode_y + barcode_h - 7, printed_at[:16])
    bar_x, bar_w = x + 6, body_w - 12
    bar_h = max(12, barcode_h - 19)
    draw_barcode(c, scan_code, bar_x, barcode_y + 7, bar_w, bar_h, min_bar_width=0.5, max_bar_width=1.6)
    size = fit_font_size(scan_code, BOLD, bar_w, 7.5, 6.2)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, size)
    c.drawCentredString(x + body_w / 2, barcode_y + 1.5, clip_text(scan_code, bar_w, BOLD, size))


def draw_opp_label(
    c,
    width: float,
    height: float,
    row: dict[str, Any],
    palette: dict[str, Any],
    logo_path: Path | None,
    printed_at: str,
    *,
    draw_logo: Callable[..., None],
    draw_barcode: Callable[..., None],
) -> None:
    _page(c, width, height)
    outer, pad = 8, 10
    c.setFillColor(colors.white)
    c.setStrokeColor(palette["line_strong"])
    c.roundRect(outer, outer, width - 2 * outer, height - 2 * outer, 6, stroke=1, fill=1)
    c.setFillColor(palette["primary"])
    c.rect(outer, height - outer - 3, width - 2 * outer, 3, stroke=0, fill=1)
    x, body_w = outer + pad, width - 2 * (outer + pad)

    header_y, header_h = height - outer - 54, 42
    _box(c, x, header_y, body_w, header_h, palette)
    draw_logo(c, palette, logo_path, x + 8, header_y + 9, 58, 24, radius=6, padding_x=3, padding_y=2, line_width=0.7)
    next_w = 100
    _label_value(c, x + body_w - next_w - 8, header_y + 8, next_w, 26, "Operacao seguinte", _text(row.get("proximo_posto")), palette, value_size=8)
    title_x, title_w = x + 78, body_w - 78 - next_w - 18
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, fit_font_size("Etiqueta OPP", BOLD, title_w, 8, 6))
    c.drawString(title_x, header_y + 23, "Etiqueta OPP")
    subtitle = f"{_text(row.get('encomenda'))} | {_text(row.get('cliente'))}"
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 6.4)
    c.drawString(title_x, header_y + 10, clip_text(subtitle, title_w, REGULAR, 6.4))

    hero_y, hero_h = header_y - 61, 52
    opp_w, gap = 116, 8
    _label_value(c, x, hero_y, body_w - opp_w - gap, hero_h, "Referencia interna", _text(row.get("ref_interna")), palette, value_size=8)
    opp_text = f"{_text(row.get('opp'))} | Qtd {_text(row.get('quantidade_txt'))}"
    _label_value(c, x + body_w - opp_w, hero_y, opp_w, hero_h, "OPP / quantidade", opp_text, palette, value_size=8, fill=palette["surface_alt"])

    details_y, details_h = hero_y - 49, 40
    _box(c, x, details_y, body_w, details_h, palette)
    desc_w = body_w * 0.54
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 5.8)
    c.drawString(x + 8, details_y + 28, "Descricao / referencia externa")
    description = _text(row.get("descricao") or row.get("ref_externa"))
    lines = wrap_text(description, REGULAR, 6.8, desc_w - 16, max_lines=2) or ["-"]
    c.setFillColor(palette["ink"])
    c.setFont(REGULAR, 6.8)
    for index, line in enumerate(lines[:2]):
        c.drawString(x + 8, details_y + 17 - 8 * index, clip_text(line, desc_w - 16, REGULAR, 6.8))
    production = f"OF {_text(row.get('of'))} | {_text(row.get('material'))} {_text(row.get('espessura'))} mm | {_text(row.get('estado'))}"
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 5.8)
    c.drawString(x + desc_w + 8, details_y + 28, "Dados de producao")
    prod_size = fit_font_size(production, REGULAR, body_w - desc_w - 24, 6.8, 5.2)
    c.setFillColor(palette["ink"])
    c.setFont(REGULAR, prod_size)
    c.drawString(x + desc_w + 8, details_y + 14, clip_text(production, body_w - desc_w - 24, REGULAR, prod_size))

    bottom_y, bottom_h = outer + 10, details_y - outer - 18
    route_w = body_w * 0.34
    _label_value(c, x, bottom_y, route_w, bottom_h, "Fluxo operacional", f"{_text(row.get('posto_origem'))} > {_text(row.get('proximo_posto'))}", palette, value_size=8, fill=palette["surface_alt"])
    barcode_x, barcode_w = x + route_w + gap, body_w - route_w - gap
    _box(c, barcode_x, bottom_y, barcode_w, bottom_h, palette)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 5.8)
    c.drawString(barcode_x + 8, bottom_y + bottom_h - 11, "Codigo OPP para picagem")
    opp = _text(row.get("opp"))
    draw_barcode(c, opp, barcode_x + 10, bottom_y + 16, barcode_w - 20, max(25, bottom_h - 38), min_bar_width=0.48, max_bar_width=1.4)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, 7)
    c.drawCentredString(barcode_x + barcode_w / 2, bottom_y + 5, clip_text(opp, barcode_w - 20, BOLD, 7))


def draw_pallet_label(
    c,
    width: float,
    height: float,
    rows: list[dict[str, Any]],
    all_rows: list[dict[str, Any]],
    destination: str,
    palette: dict[str, Any],
    logo_path: Path | None,
    source: str,
    printed_at: str,
    page_no: int,
    total_pages: int,
    *,
    draw_logo: Callable[..., None],
    fmt: Callable[[Any], str],
) -> None:
    _page(c, width, height)
    margin = 24
    c.setFillColor(palette["primary"])
    c.rect(margin, height - margin - 3, width - 2 * margin, 3, stroke=0, fill=1)
    header_y, header_h = height - 92, 58
    draw_logo(c, palette, logo_path, margin, header_y + 8, 106, 42, radius=7, padding_x=5, padding_y=3, line_width=0.7)
    title_x = margin + 126
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, 8)
    c.drawString(title_x, header_y + 32, "Etiqueta de Palete")
    subtitle = f"Destino {destination or '-'} | Origem {source or '-'}"
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 7.2)
    c.drawString(title_x, header_y + 16, clip_text(subtitle, 330, REGULAR, 7.2))
    meta_x = width - margin - 250
    meta_w = 118
    document = f"PLT-{_text(rows[0].get('encomenda') if rows else '-') }"
    _label_value(c, meta_x, header_y + 8, meta_w, 42, "Documento", document, palette, value_size=8)
    _label_value(c, meta_x + meta_w + 8, header_y + 8, meta_w, 42, "Pagina / impresso", f"{page_no}/{total_pages} | {printed_at[:16]}", palette, value_size=7.6)

    cards_y, gap = header_y - 55, 8
    card_w = (width - 2 * margin - 2 * gap) / 3
    client = _text(rows[0].get("cliente_label") if rows else "-")
    qty = sum(float(row.get("quantidade", 0) or 0) for row in all_rows)
    next_ops = ", ".join(dict.fromkeys(_text(row.get("proximo_posto")) for row in all_rows))
    for index, (label, value) in enumerate((("Cliente", client), ("Conteudo", f"Refs {len(all_rows)} | Qtd {fmt(qty)}"), ("Operacao seguinte", next_ops))):
        _label_value(c, margin + index * (card_w + gap), cards_y, card_w, 46, label, value, palette, value_size=8)

    destination_y = cards_y - 61
    _label_value(c, margin, destination_y, width - 2 * margin, 48, "Destino desta palete", destination or "-", palette, value_size=8, fill=palette["primary_soft"])

    columns = [("Ref. interna", 112), ("Ref. externa", 110), ("Descricao", 172), ("Material", 76), ("Esp.", 44), ("Qtd", 46), ("OPP", 94), ("Seguinte", width - 2 * margin - 654)]
    table_y = destination_y - 28
    c.setFillColor(palette["accent_soft"] if "accent_soft" in palette else palette["primary_soft"])
    c.setStrokeColor(palette["line"])
    c.roundRect(margin, table_y, width - 2 * margin, 20, 5, stroke=1, fill=1)
    cursor = margin
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, 7.2)
    for label, col_w in columns:
        c.drawString(cursor + 6, table_y + 7, clip_text(label, col_w - 12, BOLD, 7.2))
        cursor += col_w
    for index, row in enumerate(rows):
        y = table_y - 22 * (index + 1)
        _box(c, margin, y, width - 2 * margin, 19, palette, fill=palette["surface"] if index % 2 == 0 else colors.white, radius=4)
        values = [row.get("ref_interna"), row.get("ref_externa"), row.get("descricao"), row.get("material"), row.get("espessura"), row.get("quantidade_txt"), row.get("opp"), row.get("proximo_posto")]
        cursor = margin
        for (_, col_w), value in zip(columns, values):
            item = _text(value)
            size = fit_font_size(item, REGULAR, col_w - 12, 7, 5.5)
            c.setFillColor(palette["ink"])
            c.setFont(REGULAR, size)
            c.drawString(cursor + 6, y + 6, clip_text(item, col_w - 12, REGULAR, size))
            cursor += col_w
    c.setStrokeColor(palette["line"])
    c.line(margin, 28, width - margin, 28)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 6.8)
    c.drawString(margin, 16, f"Origem {source or '-'} | Encomenda {_text(rows[0].get('encomenda') if rows else '-')}")
    c.drawRightString(width - margin, 16, f"LUGEST | {printed_at}")
