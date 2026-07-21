from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as pdf_canvas

from .text import clip_text, fit_font_size, mix_hex


REGULAR = "Helvetica"
BOLD = "Helvetica-Bold"


def _number(value: Any, default: float = 0.0) -> float:
    try:
        return float(str(value or "0").strip().replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return float(default)


def _fmt(value: Any, decimals: int = 2) -> str:
    number = _number(value)
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.{decimals}f}".rstrip("0").rstrip(".")


def _primary(value: Any) -> str:
    text = str(value or "").strip().upper()
    if len(text) == 7 and text.startswith("#"):
        try:
            int(text[1:], 16)
            return text
        except ValueError:
            pass
    return "#00A6A6"


def _palette(primary_hex: str) -> dict[str, Any]:
    return {
        "accent": colors.HexColor(primary_hex),
        "accent_soft": colors.HexColor(mix_hex(primary_hex, "#FFFFFF", 0.88)),
        "accent_faint": colors.HexColor(mix_hex(primary_hex, "#FFFFFF", 0.95)),
        "ink": colors.HexColor("#102538"),
        "muted": colors.HexColor("#5D6F7E"),
        "line": colors.HexColor("#CBD5DD"),
        "surface": colors.HexColor("#F6F8FA"),
        "danger": colors.HexColor("#B42318"),
        "warning": colors.HexColor("#A15C00"),
        "white": colors.white,
    }


def _draw_logo(c, path: str, x: float, y: float, width: float, height: float, palette: dict[str, Any]) -> None:
    c.setFillColor(palette["white"])
    c.setStrokeColor(palette["line"])
    c.roundRect(x, y, width, height, 7, stroke=1, fill=1)
    logo = Path(str(path or ""))
    if not logo.exists():
        c.setFillColor(palette["ink"])
        c.setFont(BOLD, 8)
        c.drawCentredString(x + width / 2, y + height / 2 - 3, "LUGEST")
        return
    try:
        image = ImageReader(str(logo))
        iw, ih = image.getSize()
        scale = min((width - 12) / max(iw, 1), (height - 10) / max(ih, 1))
        draw_w, draw_h = iw * scale, ih * scale
        c.drawImage(image, x + (width - draw_w) / 2, y + (height - draw_h) / 2, draw_w, draw_h, mask="auto")
    except Exception:
        c.setFillColor(palette["ink"])
        c.setFont(BOLD, 8)
        c.drawCentredString(x + width / 2, y + height / 2 - 3, "LUGEST")


def _card(c, x: float, y: float, width: float, height: float, palette: dict[str, Any], *, fill=None, radius: float = 3) -> None:
    c.setFillColor(fill or palette["white"])
    c.setStrokeColor(palette["line"])
    c.setLineWidth(0.75)
    c.roundRect(x, y, width, height, radius, stroke=1, fill=1)


def _metric(c, x: float, y: float, width: float, label: str, value: str, palette: dict[str, Any], *, tone: str = "normal") -> None:
    fill = palette["accent_faint"] if tone == "accent" else palette["white"]
    _card(c, x, y, width, 38, palette, fill=fill)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 6.5)
    c.drawString(x + 8, y + 24, clip_text(label, width - 16, REGULAR, 6.5))
    color = palette["danger"] if tone == "danger" else palette["ink"]
    c.setFillColor(color)
    size = fit_font_size(value, BOLD, width - 16, 8, 6)
    c.setFont(BOLD, size)
    c.drawString(x + 8, y + 8, clip_text(value, width - 16, BOLD, size))


def _header(
    c,
    page_w: float,
    page_h: float,
    branding: dict[str, Any],
    palette: dict[str, Any],
    title: str,
    subtitle: str,
    page_no: int,
    total_pages: int,
) -> None:
    margin = 24
    top = page_h - margin
    c.setFillColor(palette["accent"])
    c.rect(margin, top - 3, page_w - 2 * margin, 3, stroke=0, fill=1)
    _draw_logo(c, str(branding.get("logo_path", "")), margin, top - 63, 118, 48, palette)
    title_x = margin + 136
    meta_w = 258
    meta_x = page_w - margin - meta_w
    available_w = meta_x - title_x - 20
    title_size = fit_font_size(title, BOLD, available_w, 8, 6.2)
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, title_size)
    c.drawString(title_x, top - 30, clip_text(title, available_w, BOLD, title_size))
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 7.2)
    c.drawString(title_x, top - 47, clip_text(subtitle, available_w, REGULAR, 7.2))

    gap = 7
    chip_w = (meta_w - 2 * gap) / 3
    meta = (("Emitido", datetime.now().strftime("%d/%m/%Y")), ("Documento", "Inventario"), ("Pagina", f"{page_no}/{total_pages}"))
    for index, (label, value) in enumerate(meta):
        x = meta_x + index * (chip_w + gap)
        _card(c, x, top - 62, chip_w, 40, palette)
        c.setFillColor(palette["muted"])
        c.setFont(REGULAR, 6.1)
        c.drawString(x + 7, top - 37, label)
        size = fit_font_size(value, BOLD, chip_w - 14, 8, 6)
        c.setFillColor(palette["ink"])
        c.setFont(BOLD, size)
        c.drawString(x + 7, top - 53, clip_text(value, chip_w - 14, BOLD, size))


def _table_header(c, x: float, y: float, columns: list[tuple[str, float, str]], palette: dict[str, Any]) -> None:
    width = sum(item[1] for item in columns)
    c.setFillColor(palette["accent_soft"])
    c.setStrokeColor(palette["line"])
    c.rect(x, y, width, 21, stroke=1, fill=1)
    cursor = x
    c.setFillColor(palette["ink"])
    c.setFont(BOLD, 7.2)
    for label, col_w, align in columns:
        if align == "right":
            c.drawRightString(cursor + col_w - 6, y + 7, label)
        elif align == "center":
            c.drawCentredString(cursor + col_w / 2, y + 7, label)
        else:
            c.drawString(cursor + 6, y + 7, label)
        cursor += col_w


def _table_row(c, x: float, y: float, columns, values, palette, index: int, status: str = "") -> None:
    width = sum(item[1] for item in columns)
    fill = palette["surface"] if index % 2 == 0 else palette["white"]
    c.setFillColor(fill)
    c.setStrokeColor(palette["line"])
    c.setLineWidth(0.4)
    c.rect(x, y, width, 18, stroke=1, fill=1)
    color = palette["danger"] if status == "danger" else palette["warning"] if status == "warning" else palette["ink"]
    c.setFillColor(color)
    cursor = x
    for (label, col_w, align), value in zip(columns, values):
        text = str(value if value not in (None, "") else "-")
        size = fit_font_size(text, REGULAR, col_w - 12, 7.3, 5.8)
        c.setFont(REGULAR, size)
        shown = clip_text(text, col_w - 12, REGULAR, size)
        if align == "right":
            c.drawRightString(cursor + col_w - 6, y + 6.1, shown)
        elif align == "center":
            c.drawCentredString(cursor + col_w / 2, y + 6.1, shown)
        else:
            c.drawString(cursor + 6, y + 6.1, shown)
        cursor += col_w


def _footer(c, page_w: float, page_no: int, total_pages: int, palette: dict[str, Any], branding: dict[str, Any]) -> None:
    margin = 24
    lines = [str(item).strip() for item in list(branding.get("empresa_info_rodape", []) or []) if str(item).strip()]
    c.setStrokeColor(palette["line"])
    c.line(margin, 27, page_w - margin, 27)
    c.setFillColor(palette["muted"])
    c.setFont(REGULAR, 6.6)
    c.drawString(margin, 15, clip_text(" | ".join(lines[:2]) or "LUGEST", page_w - 190, REGULAR, 6.6))
    c.drawRightString(page_w - margin, 15, f"Inventario | {page_no}/{total_pages}")


def _material_dimension(row: dict[str, Any], formato: str) -> str:
    raw = str(row.get("dimensao", row.get("dimensoes", "")) or "").strip()
    comp, larg = _number(row.get("comprimento")), _number(row.get("largura"))
    esp, diam, metros = _number(row.get("espessura")), _number(row.get("diametro")), _number(row.get("metros"))
    if formato.casefold() == "chapa" and comp > 0 and larg > 0:
        return f"{_fmt(comp)} x {_fmt(larg)} mm"
    if formato.casefold() == "tubo":
        if diam > 0:
            return f"D{_fmt(diam)} x {_fmt(esp)} mm | {_fmt(metros)} m"
        if comp > 0 and larg > 0:
            return f"{_fmt(comp)} x {_fmt(larg)} x {_fmt(esp)} mm | {_fmt(metros)} m"
    if raw:
        return raw
    if comp > 0 or larg > 0:
        return f"{_fmt(comp)} x {_fmt(larg)} mm"
    return f"{_fmt(metros)} m" if metros > 0 else "-"


def _location(row: dict[str, Any]) -> str:
    for key in ("Localizacao", "Localização", "localizacao", "local", "LocalizaÃ§Ã£o", "LocalizaÃƒÂ§ÃƒÂ£o"):
        value = str(row.get(key, "") or "").strip()
        if value:
            return value
    return "-"


def _render_inventory(
    path: str | Path,
    *,
    title: str,
    subtitle: str,
    rows: list[dict[str, Any]],
    columns: list[tuple[str, float, str]],
    values_for: Callable[[dict[str, Any]], list[Any]],
    status_for: Callable[[dict[str, Any]], str],
    metrics: list[tuple[str, str, str]],
    branding: dict[str, Any],
) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_w, page_h = landscape(A4)
    c = pdf_canvas.Canvas(str(target), pagesize=(page_w, page_h))
    palette = _palette(_primary(branding.get("primary_color")))
    margin, row_h = 24, 18
    table_top = page_h - 158
    table_bottom = 38
    rows_per_page = max(1, int((table_top - table_bottom - 21) // row_h))
    total_pages = max(1, math.ceil(len(rows) / rows_per_page))

    for page_index in range(total_pages):
        page_no = page_index + 1
        c.setFillColor(colors.white)
        c.rect(0, 0, page_w, page_h, stroke=0, fill=1)
        _header(c, page_w, page_h, branding, palette, title, subtitle, page_no, total_pages)
        gap = 8
        metric_w = (page_w - 2 * margin - gap * 3) / 4
        metric_y = page_h - 132
        for index, (label, value, tone) in enumerate(metrics[:4]):
            _metric(c, margin + index * (metric_w + gap), metric_y, metric_w, label, value, palette, tone=tone)
        header_y = table_top - 21
        _table_header(c, margin, header_y, columns, palette)
        page_rows = rows[page_index * rows_per_page : (page_index + 1) * rows_per_page]
        for local_index, row in enumerate(page_rows):
            row_y = header_y - (local_index + 1) * row_h
            _table_row(c, margin, row_y, columns, values_for(row), palette, local_index, status_for(row))
        if not page_rows:
            c.setFillColor(palette["muted"])
            c.setFont(REGULAR, 8)
            c.drawCentredString(page_w / 2, header_y - 36, "Sem registos para apresentar.")
        _footer(c, page_w, page_no, total_pages, palette, branding)
        if page_no < total_pages:
            c.showPage()
    c.save()
    return target


def render_material_stock_pdf(path: str | Path, data: dict[str, Any], branding: dict[str, Any], *, in_stock_only: bool = False) -> Path:
    rows = [dict(item) for item in list(data.get("materiais", []) or []) if isinstance(item, dict)]
    if in_stock_only:
        rows = [row for row in rows if _number(row.get("quantidade")) - _number(row.get("reservado")) > 0]
    rows.sort(key=lambda row: (str(row.get("formato", "")), str(row.get("material", "")), _number(row.get("espessura")), str(row.get("id", ""))))
    total = sum(_number(row.get("quantidade")) for row in rows)
    reserved = sum(_number(row.get("reservado")) for row in rows)
    available = sum(max(0.0, _number(row.get("quantidade")) - _number(row.get("reservado"))) for row in rows)
    retalhos = sum(1 for row in rows if bool(row.get("is_sobra")) or "retalho" in _location(row).casefold())
    formats = len({str(row.get("formato", "") or "Chapa").strip() for row in rows})
    columns = [
        ("ID", 76, "left"), ("Formato", 62, "left"), ("Material", 88, "left"),
        ("Dimensao / secao", 196, "left"), ("Esp.", 42, "center"),
        ("Stock", 48, "right"), ("Reserv.", 48, "right"), ("Disp.", 48, "right"),
        ("Localizacao", 112, "left"), ("Lote", 73, "left"),
    ]

    def values(row: dict[str, Any]) -> list[Any]:
        formato = str(row.get("formato", "") or "Chapa").strip() or "Chapa"
        qty, res = _number(row.get("quantidade")), _number(row.get("reservado"))
        return [row.get("id"), formato, row.get("material"), _material_dimension(row, formato), _fmt(row.get("espessura")), _fmt(qty), _fmt(res), _fmt(max(0, qty - res)), _location(row), row.get("lote_fornecedor", row.get("lote", "-"))]

    def status(row: dict[str, Any]) -> str:
        available_qty = _number(row.get("quantidade")) - _number(row.get("reservado"))
        if available_qty <= 0:
            return "danger"
        if bool(row.get("is_sobra")) or "retalho" in _location(row).casefold():
            return "warning"
        return ""

    return _render_inventory(
        path, title="Stock de Materia-Prima", subtitle="Disponibilidade fisica, reserva e rastreabilidade por lote",
        rows=rows, columns=columns, values_for=values, status_for=status,
        metrics=[("Referencias", str(len(rows)), "normal"), ("Stock fisico", _fmt(total), "accent"), ("Disponivel", _fmt(available), "accent"), ("Reservado | Retalhos | Formatos", f"{_fmt(reserved)} | {retalhos} | {formats}", "normal")],
        branding=branding,
    )


def render_product_stock_pdf(
    path: str | Path,
    data: dict[str, Any],
    branding: dict[str, Any],
    unit_price: Callable[[dict[str, Any]], float],
) -> Path:
    rows = [dict(item) for item in list(data.get("produtos", []) or []) if isinstance(item, dict)]
    rows.sort(key=lambda row: (str(row.get("categoria", "")), str(row.get("descricao", "")), str(row.get("codigo", ""))))
    total_qty = sum(_number(row.get("qty")) for row in rows)
    total_value = sum(_number(row.get("qty")) * _number(unit_price(row)) for row in rows)
    alerts = sum(1 for row in rows if _number(row.get("qty")) <= 0 or (_number(row.get("alerta")) > 0 and _number(row.get("qty")) <= _number(row.get("alerta"))))
    categories = len({str(row.get("categoria", "") or "Sem categoria") for row in rows})
    columns = [
        ("Codigo", 78, "left"), ("Descricao", 214, "left"), ("Categoria", 96, "left"),
        ("Tipo / dimensao", 128, "left"), ("Un.", 35, "center"), ("Stock", 48, "right"),
        ("Alerta", 48, "right"), ("Preco/un.", 66, "right"), ("Valor", 80, "right"),
    ]

    def product_dims(row: dict[str, Any]) -> str:
        explicit = str(row.get("dimensoes", "") or "").strip()
        if explicit:
            return explicit
        dims = [_number(row.get(key)) for key in ("comprimento", "largura", "espessura")]
        return " x ".join(_fmt(value) for value in dims) if any(dims) else "-"

    def values(row: dict[str, Any]) -> list[Any]:
        qty = _number(row.get("qty"))
        price = _number(unit_price(row))
        kind = str(row.get("tipo", "") or "-").strip()
        dimension = product_dims(row)
        return [row.get("codigo"), row.get("descricao"), row.get("categoria"), f"{kind} | {dimension}", row.get("unid", "UN"), _fmt(qty), _fmt(row.get("alerta")), f"{price:.2f}", f"{qty * price:.2f}"]

    def status(row: dict[str, Any]) -> str:
        qty, alert = _number(row.get("qty")), _number(row.get("alerta"))
        return "danger" if qty <= 0 else "warning" if alert > 0 and qty <= alert else ""

    return _render_inventory(
        path, title="Stock de Produtos", subtitle="Existencias, niveis de reposicao e valorizacao atualizada",
        rows=rows, columns=columns, values_for=values, status_for=status,
        metrics=[("Referencias", str(len(rows)), "normal"), ("Unidades em stock", _fmt(total_qty), "accent"), ("Reposicao necessaria", str(alerts), "danger" if alerts else "normal"), ("Categorias | Valor global", f"{categories} | {total_value:.2f} EUR", "accent")],
        branding=branding,
    )
