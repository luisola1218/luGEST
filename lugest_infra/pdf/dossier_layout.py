from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from reportlab.lib import colors
from reportlab.lib.units import mm

from .text import clip_text, fit_font_size, mix_hex, wrap_text


REGULAR = "Helvetica"
BOLD = "Helvetica-Bold"


def _hex(value: Any) -> str:
    text = str(value or "").strip().upper()
    if len(text) == 7 and text.startswith("#"):
        try:
            int(text[1:], 16)
            return text
        except ValueError:
            pass
    return "#00A6A6"


def dossier_palette(primary: Any) -> dict[str, Any]:
    primary_hex = _hex(primary)
    return {
        "primary": colors.HexColor(primary_hex),
        "primary_soft": colors.HexColor(mix_hex(primary_hex, "#FFFFFF", 0.88)),
        "primary_faint": colors.HexColor(mix_hex(primary_hex, "#FFFFFF", 0.95)),
        "navy": colors.HexColor("#0B1F33"),
        "steel": colors.HexColor("#34495E"),
        "ink": colors.HexColor("#14212B"),
        "muted": colors.HexColor("#61717F"),
        "line": colors.HexColor("#CAD3DA"),
        "line_soft": colors.HexColor("#DFE5EA"),
        "surface": colors.HexColor("#F3F6F8"),
        "white": colors.white,
        "green": colors.HexColor("#78BE20"),
        "amber": colors.HexColor("#F0A202"),
        "warning_fill": colors.HexColor("#FFF8EB"),
        "danger": colors.HexColor("#B42318"),
        "danger_fill": colors.HexColor("#FFF1F2"),
        "success": colors.HexColor("#107569"),
        "success_fill": colors.HexColor("#ECFDF3"),
    }


@dataclass(frozen=True)
class Column:
    label: str
    width: float
    align: str = "left"


class DossierLayout:
    """Shared, bounded PDF geometry based on the technical dossier."""

    def __init__(
        self,
        canvas_obj,
        page_size: tuple[float, float],
        *,
        primary: Any = "#00A6A6",
        logo_path: Path | None = None,
        logo_draw: Callable[..., None] | None = None,
        issued_at: str = "",
        document_code: str = "",
        revision: str = "REV. 00",
        margin: float | None = None,
    ) -> None:
        self.c = canvas_obj
        self.width, self.height = page_size
        self.margin = float(margin if margin is not None else 11 * mm)
        self.inner_width = self.width - 2 * self.margin
        self.palette = dossier_palette(primary)
        self.logo_path = logo_path
        self.logo_draw = logo_draw
        self.issued_at = str(issued_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.document_code = str(document_code or "DOCUMENTO").strip() or "DOCUMENTO"
        self.revision = str(revision or "REV. 00").strip() or "REV. 00"

    def _logo(self, x: float, y: float, width: float, height: float) -> None:
        if callable(self.logo_draw):
            try:
                self.logo_draw(
                    self.c,
                    self.palette,
                    self.logo_path,
                    x,
                    y,
                    width,
                    height,
                    radius=3,
                    padding_x=3,
                    padding_y=2,
                    line_width=0.7,
                )
                return
            except Exception:
                pass
        self.c.setFillColor(self.palette["white"])
        self.c.setStrokeColor(self.palette["line"])
        self.c.roundRect(x, y, width, height, 3, stroke=1, fill=1)
        self.c.setFillColor(self.palette["navy"])
        self.c.setFont(BOLD, 8)
        self.c.drawCentredString(x + width / 2, y + height / 2 - 3, "LUGEST")

    def begin_page(
        self,
        title: str,
        subtitle: str,
        page_no: int,
        total_pages: int,
        *,
        status: str = "",
        section_label: str = "DOCUMENTO TECNICO",
    ) -> float:
        c, p = self.c, self.palette
        c.setFillColor(p["white"])
        c.rect(0, 0, self.width, self.height, stroke=0, fill=1)
        c.setFillColor(p["primary"])
        c.rect(0, self.height - 3 * mm, self.width, 3 * mm, stroke=0, fill=1)
        top = self.height - self.margin
        header_h = 27 * mm
        c.setFillColor(p["white"])
        c.setStrokeColor(p["line"])
        c.roundRect(self.margin, top - header_h, self.inner_width, header_h, 4, stroke=1, fill=1)
        c.setFillColor(p["primary"])
        c.rect(self.margin, top - header_h, 3.2 * mm, header_h, stroke=0, fill=1)
        logo_w, logo_h = 37 * mm, 16 * mm
        self._logo(self.margin + 7 * mm, top - 21.5 * mm, logo_w, logo_h)

        title_x = self.margin + 49 * mm
        meta_w = 46 * mm
        title_w = self.inner_width - 49 * mm - meta_w - 7 * mm
        c.setFillColor(p["steel"])
        c.setFont(BOLD, 6.1)
        c.drawString(title_x, top - 7.0 * mm, clip_text(section_label.upper(), title_w, BOLD, 6.1))
        size = fit_font_size(title, BOLD, title_w, 8.0, 6.2)
        title_lines = wrap_text(title, BOLD, size, title_w, max_lines=2) or [str(title or "DOCUMENTO")]
        c.setFillColor(p["navy"])
        c.setFont(BOLD, size)
        title_y = top - (12.5 * mm if len(title_lines) == 1 else 10.8 * mm)
        for line in title_lines[:2]:
            c.drawString(title_x, title_y, clip_text(line, title_w, BOLD, size))
            title_y -= 5.3 * mm
        c.setFillColor(p["muted"])
        c.setFont(REGULAR, 6.4)
        c.drawString(title_x, top - 23.0 * mm, clip_text(subtitle, title_w, REGULAR, 6.4))

        meta_x = self.width - self.margin - meta_w - 4 * mm
        self.label_value(meta_x, top - 4 * mm, 21 * mm, "PAGINA", f"{page_no}/{total_pages}", height=9 * mm, compact=True)
        self.label_value(meta_x + 23 * mm, top - 4 * mm, 23 * mm, "REVISAO", self.revision, height=9 * mm, compact=True)
        self.label_value(meta_x, top - 15 * mm, 46 * mm, "DOCUMENTO", self.document_code, height=9 * mm, compact=True)
        if status:
            c.setFillColor(p["green"])
            c.roundRect(meta_x, top - 27 * mm, 46 * mm, 6.5 * mm, 2, stroke=0, fill=1)
            c.setFillColor(p["navy"])
            c.setFont(BOLD, 5.8)
            c.drawCentredString(meta_x + 23 * mm, top - 24.8 * mm, clip_text(status.upper(), 42 * mm, BOLD, 5.8))
        return top - header_h - 5 * mm

    def footer(self, page_no: int, total_pages: int, section: str = "") -> None:
        c, p = self.c, self.palette
        y = self.margin + 5.2 * mm
        c.setStrokeColor(p["line"])
        c.line(self.margin, y, self.width - self.margin, y)
        c.setFillColor(p["muted"])
        c.setFont(REGULAR, 5.8)
        c.drawString(self.margin, self.margin + 1.8 * mm, clip_text(f"DOCUMENTO CONTROLADO | {self.document_code} | {self.revision}", self.inner_width * 0.42, REGULAR, 5.8))
        c.drawCentredString(self.width / 2, self.margin + 1.8 * mm, clip_text(section, self.inner_width * 0.30, REGULAR, 5.8))
        c.drawRightString(self.width - self.margin, self.margin + 1.8 * mm, f"{page_no}/{total_pages} | {self.issued_at[:19]}")

    def section(self, y_top: float, index: str, title: str, subtitle: str = "") -> float:
        c, p = self.c, self.palette
        h = 8 * mm
        c.setFillColor(p["surface"])
        c.setStrokeColor(p["line"])
        c.rect(self.margin, y_top - h, self.inner_width, h, stroke=1, fill=1)
        c.setFillColor(p["primary"])
        c.rect(self.margin, y_top - h, 13 * mm, h, stroke=0, fill=1)
        c.setFillColor(p["white"])
        c.setFont(BOLD, 7.2)
        c.drawCentredString(self.margin + 6.5 * mm, y_top - 5.4 * mm, str(index))
        c.setFillColor(p["navy"])
        c.setFont(BOLD, 8)
        c.drawString(self.margin + 17 * mm, y_top - 5.4 * mm, clip_text(str(title).upper(), self.inner_width * 0.58, BOLD, 8))
        if subtitle:
            c.setFillColor(p["muted"])
            c.setFont(REGULAR, 5.8)
            c.drawRightString(self.width - self.margin - 3 * mm, y_top - 5.2 * mm, clip_text(subtitle, self.inner_width * 0.34, REGULAR, 5.8))
        return y_top - 11 * mm

    def card(self, x: float, y_top: float, width: float, height: float, *, fill=None, border=None, radius: float = 3) -> None:
        self.c.setFillColor(fill or self.palette["white"])
        self.c.setStrokeColor(border or self.palette["line"])
        self.c.setLineWidth(0.75)
        self.c.roundRect(x, y_top - height, width, height, radius, stroke=1, fill=1)

    def label_value(
        self,
        x: float,
        y_top: float,
        width: float,
        label: str,
        value: str,
        *,
        height: float = 15 * mm,
        fill=None,
        compact: bool = False,
        max_lines: int = 2,
    ) -> None:
        c, p = self.c, self.palette
        self.card(x, y_top, width, height, fill=fill or p["surface"])
        label_size = 4.8 if compact else 6.0
        c.setFillColor(p["steel"])
        c.setFont(BOLD, label_size)
        c.drawString(x + 2.5 * mm, y_top - (3.2 if compact else 4.6) * mm, clip_text(str(label).upper(), width - 5 * mm, BOLD, label_size))
        value_size = 6.2 if compact else 7.4
        lines = wrap_text(str(value or "-"), REGULAR, value_size, width - 5 * mm, max_lines=max_lines) or ["-"]
        c.setFillColor(p["ink"])
        c.setFont(REGULAR, value_size)
        cursor = y_top - (6.5 if compact else 9.0) * mm
        line_step = 2.8 * mm if compact else 3.6 * mm
        bottom = y_top - height + 2.2 * mm
        for line in lines:
            if cursor < bottom:
                break
            c.drawString(x + 2.5 * mm, cursor, clip_text(line, width - 5 * mm, REGULAR, value_size))
            cursor -= line_step

    def metrics(self, y_top: float, entries: Iterable[tuple[str, str, Any]], *, gap: float = 3 * mm, height: float = 15 * mm) -> float:
        rows = list(entries)
        if not rows:
            return y_top
        width = (self.inner_width - gap * (len(rows) - 1)) / len(rows)
        for index, (label, value, accent) in enumerate(rows):
            x = self.margin + index * (width + gap)
            self.card(x, y_top, width, height)
            self.c.setFillColor(accent or self.palette["primary"])
            self.c.rect(x, y_top - height, 3 * mm, height, stroke=0, fill=1)
            self.c.setFillColor(self.palette["muted"])
            self.c.setFont(BOLD, 5.8)
            self.c.drawString(x + 5 * mm, y_top - 5.3 * mm, clip_text(label.upper(), width - 7 * mm, BOLD, 5.8))
            size = fit_font_size(str(value), BOLD, width - 7 * mm, 8.0, 6.0)
            self.c.setFillColor(self.palette["ink"])
            self.c.setFont(BOLD, size)
            self.c.drawString(x + 5 * mm, y_top - 11.8 * mm, clip_text(str(value), width - 7 * mm, BOLD, size))
        return y_top - height - 4 * mm

    def table_header(self, y_top: float, columns: list[Column], *, height: float = 8 * mm) -> float:
        c, p = self.c, self.palette
        total = sum(column.width for column in columns)
        c.setFillColor(p["surface"])
        c.setStrokeColor(p["line"])
        c.rect(self.margin, y_top - height, total, height, stroke=1, fill=1)
        c.setFillColor(p["navy"])
        c.setFont(BOLD, 5.8)
        x = self.margin
        for column in columns:
            label = clip_text(column.label, column.width - 4, BOLD, 5.8)
            if column.align == "right":
                c.drawRightString(x + column.width - 2, y_top - 5.2 * mm, label)
            elif column.align == "center":
                c.drawCentredString(x + column.width / 2, y_top - 5.2 * mm, label)
            else:
                c.drawString(x + 2, y_top - 5.2 * mm, label)
            x += column.width
        return y_top - height

    def table_row(
        self,
        y_top: float,
        columns: list[Column],
        values: list[Any],
        *,
        height: float = 9 * mm,
        index: int = 0,
        tone: str = "",
        font_size: float = 6.0,
        max_lines: int = 2,
    ) -> float:
        c, p = self.c, self.palette
        fill = p["white"] if index % 2 == 0 else p["surface"]
        text_color = p["ink"]
        if tone == "danger":
            fill, text_color = p["danger_fill"], p["danger"]
        elif tone == "warning":
            fill = p["warning_fill"]
        elif tone == "success":
            fill = p["success_fill"]
        total = sum(column.width for column in columns)
        c.setFillColor(fill)
        c.setStrokeColor(p["line"])
        c.rect(self.margin, y_top - height, total, height, stroke=1, fill=1)
        x = self.margin
        for col_index, (column, raw) in enumerate(zip(columns, values)):
            value = str(raw if raw not in (None, "") else "-")
            font = BOLD if col_index == 0 else REGULAR
            size = fit_font_size(value, font, column.width - 5, font_size, 4.8)
            lines = wrap_text(value, font, size, column.width - 5, max_lines=max_lines) or ["-"]
            c.setFillColor(text_color)
            c.setFont(font, size)
            if column.align == "right":
                c.drawRightString(x + column.width - 2, y_top - height / 2 - size / 3, clip_text(lines[0], column.width - 5, font, size))
            elif column.align == "center":
                c.drawCentredString(x + column.width / 2, y_top - height / 2 - size / 3, clip_text(lines[0], column.width - 5, font, size))
            else:
                cursor = y_top - 3.2 * mm
                for line in lines:
                    if cursor < y_top - height + 2 * mm:
                        break
                    c.drawString(x + 2.5, cursor, clip_text(line, column.width - 5, font, size))
                    cursor -= 2.8 * mm
            x += column.width
        return y_top - height
