from __future__ import annotations

import copy
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pypdf import PdfReader
from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import code128
from reportlab.graphics.barcode import qr as rl_qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A5, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdf_canvas

from lugest_infra.pdf.text import clip_text, wrap_text
from lugest_qt.services.legacy_backend import LegacyBackend


NAVY = colors.HexColor("#0B1F33")
STEEL = colors.HexColor("#34495E")
TEAL = colors.HexColor("#00A6A6")
GREEN = colors.HexColor("#78BE20")
AMBER = colors.HexColor("#F0A202")
RED = colors.HexColor("#D64545")
INK = colors.HexColor("#14212B")
MUTED = colors.HexColor("#61717F")
LINE = colors.HexColor("#CAD3DA")
SURFACE = colors.HexColor("#F3F6F8")
WHITE = colors.white


def _fonts() -> tuple[str, str]:
    regular, bold = "Helvetica", "Helvetica-Bold"
    for name, filename in (("SegoeUI", "segoeui.ttf"), ("SegoeUI-Bold", "segoeuib.ttf")):
        path = Path(r"C:\Windows\Fonts") / filename
        if not path.exists():
            continue
        try:
            pdfmetrics.registerFont(TTFont(name, str(path)))
            if "Bold" in name:
                bold = name
            else:
                regular = name
        except Exception:
            pass
    return regular, bold


REGULAR, BOLD = _fonts()


def _txt(value: Any) -> str:
    return str(value or "").replace("\r", " ").replace("\n", " ").strip()


def _fmt(value: Any, decimals: int = 2) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.{decimals}f}"


def _draw_logo(c: Any, logo_path: Path | None, x: float, y: float, width: float, height: float) -> None:
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(x, y, width, height, 3, stroke=1, fill=1)
    if logo_path:
        try:
            c.drawImage(ImageReader(str(logo_path)), x + 4, y + 4, width - 8, height - 8, preserveAspectRatio=True, mask="auto", anchor="c")
            return
        except Exception:
            pass
    c.setFillColor(NAVY)
    c.setFont(BOLD, 9)
    c.drawCentredString(x + width / 2, y + height / 2 - 3, "LUGEST")


def _header(
    c: Any,
    size: tuple[float, float],
    logo_path: Path | None,
    title: str,
    subtitle: str,
    document: str,
    status: str,
    *,
    page: str = "",
) -> float:
    width, height = size
    margin = 12 * mm
    top = height - margin
    panel_h = 27 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(margin, top - panel_h, width - 2 * margin, panel_h, 3, stroke=1, fill=1)
    c.setFillColor(TEAL)
    c.rect(margin, top - 2 * mm, width - 2 * margin, 2 * mm, stroke=0, fill=1)
    c.rect(margin, top - panel_h, 2 * mm, panel_h - 2 * mm, stroke=0, fill=1)
    _draw_logo(c, logo_path, margin + 8 * mm, top - 21 * mm, 34 * mm, 15 * mm)
    title_x = margin + 47 * mm
    title_w = width - title_x - margin - 44 * mm
    c.setFillColor(NAVY)
    c.setFont(BOLD, 14)
    c.drawString(title_x, top - 9 * mm, clip_text(title, title_w, BOLD, 14))
    c.setFillColor(MUTED)
    c.setFont(REGULAR, 6.8)
    c.drawString(title_x, top - 15 * mm, clip_text(subtitle, title_w, REGULAR, 6.8))
    c.setFillColor(STEEL)
    c.setFont(BOLD, 6)
    c.drawString(title_x, top - 21 * mm, clip_text(document, title_w, BOLD, 6))
    badge_w = 37 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(GREEN)
    c.roundRect(width - margin - badge_w - 6 * mm, top - 12 * mm, badge_w, 7 * mm, 2, stroke=1, fill=1)
    c.setFillColor(GREEN)
    c.setFont(BOLD, 6)
    c.drawCentredString(width - margin - badge_w / 2 - 6 * mm, top - 9.5 * mm, clip_text(status.upper(), badge_w - 4 * mm, BOLD, 6))
    if page:
        c.setFillColor(MUTED)
        c.setFont(REGULAR, 5.5)
        c.drawRightString(width - margin - 6 * mm, top - 20.8 * mm, page)
    return top - panel_h - 6 * mm


def _footer(c: Any, size: tuple[float, float], label: str, document: str, page: str = "") -> None:
    width, _height = size
    margin = 12 * mm
    c.setStrokeColor(LINE)
    c.line(margin, 11 * mm, width - margin, 11 * mm)
    c.setFillColor(MUTED)
    c.setFont(REGULAR, 5.6)
    c.drawString(margin, 7 * mm, label)
    c.drawCentredString(width / 2, 7 * mm, document)
    c.drawRightString(width - margin, 7 * mm, f"{page} | {datetime.now().strftime('%d/%m/%Y %H:%M')}".strip(" |"))


def _section(c: Any, x: float, y_top: float, width: float, title: str, code: str = "") -> float:
    height = 8 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.rect(x, y_top - height, width, height, stroke=0, fill=1)
    c.setFillColor(TEAL)
    c.rect(x, y_top - height, 2 * mm, height, stroke=0, fill=1)
    c.setStrokeColor(LINE)
    c.line(x, y_top - height, x + width, y_top - height)
    if code:
        c.setFillColor(TEAL)
        c.setFont(BOLD, 6.2)
        c.drawCentredString(x + 8 * mm, y_top - 5.3 * mm, code)
        text_x = x + 17 * mm
    else:
        text_x = x + 4 * mm
    c.setFillColor(NAVY)
    c.setFont(BOLD, 8)
    c.drawString(text_x, y_top - 5.3 * mm, title.upper())
    return y_top - height - 4 * mm


def _metric(c: Any, x: float, y_top: float, width: float, label: str, value: str, accent: Any = TEAL, height: float = 18 * mm) -> None:
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(x, y_top - height, width, height, 3, stroke=1, fill=1)
    c.setFillColor(accent)
    c.rect(x, y_top - height, 3 * mm, height, stroke=0, fill=1)
    c.setFillColor(MUTED)
    c.setFont(BOLD, 5.6)
    c.drawString(x + 5 * mm, y_top - 5 * mm, label.upper())
    c.setFillColor(INK)
    c.setFont(BOLD, 13)
    c.drawString(x + 5 * mm, y_top - 13 * mm, clip_text(value, width - 8 * mm, BOLD, 13))


def _field(c: Any, x: float, y_top: float, width: float, label: str, value: str, *, height: float = 16 * mm) -> None:
    c.setFillColor(SURFACE)
    c.setStrokeColor(LINE)
    c.roundRect(x, y_top - height, width, height, 2.5, stroke=1, fill=1)
    c.setFillColor(MUTED)
    c.setFont(BOLD, 5.4)
    c.drawString(x + 3 * mm, y_top - 4.6 * mm, label.upper())
    c.setFillColor(INK)
    c.setFont(REGULAR, 7.2)
    lines = wrap_text(value or "-", REGULAR, 7.2, width - 6 * mm, max_lines=2) or ["-"]
    line_y = y_top - 9.3 * mm
    for line in lines:
        c.drawString(x + 3 * mm, line_y, line)
        line_y -= 3.3 * mm


def _barcode(c: Any, value: str, x: float, y: float, width: float, height: float) -> None:
    try:
        barcode = code128.Code128(value, barHeight=height - 4 * mm, barWidth=0.32)
        scale = min(1.0, width / max(barcode.width, 1))
        c.saveState()
        c.translate(x + (width - barcode.width * scale) / 2, y + 4 * mm)
        c.scale(scale, 1)
        barcode.drawOn(c, 0, 0)
        c.restoreState()
    except Exception:
        pass
    c.setFillColor(INK)
    c.setFont(BOLD, 5)
    c.drawCentredString(x + width / 2, y + 1.5 * mm, clip_text(value, width, BOLD, 5))


def _qr(c: Any, value: str, x: float, y: float, size: float) -> None:
    widget = rl_qr.QrCodeWidget(value)
    bounds = widget.getBounds()
    w = max(bounds[2] - bounds[0], 1)
    h = max(bounds[3] - bounds[1], 1)
    drawing = Drawing(size, size, transform=[size / w, 0, 0, size / h, 0, 0])
    drawing.add(widget)
    renderPDF.draw(drawing, c, x, y)


def _eur(value: Any) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    raw = f"{number:,.2f}"
    return raw.replace(",", "#").replace(".", ",").replace("#", ".") + " EUR"


def _light_header(
    c: Any,
    size: tuple[float, float],
    logo_path: Path | None,
    title: str,
    number: str,
    status: str,
    page: str,
) -> float:
    width, height = size
    margin = 12 * mm
    top = height - margin
    c.setFillColor(TEAL)
    c.rect(margin, top - 2 * mm, width - 2 * margin, 2 * mm, stroke=0, fill=1)
    _draw_logo(c, logo_path, margin, top - 22 * mm, 33 * mm, 15 * mm)
    c.setFillColor(NAVY)
    c.setFont(BOLD, 15)
    c.drawString(margin + 39 * mm, top - 12 * mm, title.upper())
    c.setFillColor(MUTED)
    c.setFont(REGULAR, 6.5)
    c.drawString(margin + 39 * mm, top - 18 * mm, "Documento comercial | Emissao digital LUGEST")
    right = width - margin
    c.setFillColor(INK)
    c.setFont(BOLD, 8)
    c.drawRightString(right, top - 10 * mm, number or "DOCUMENTO EM PREPARACAO")
    c.setStrokeColor(TEAL)
    c.setFillColor(WHITE)
    c.roundRect(right - 40 * mm, top - 20 * mm, 40 * mm, 7 * mm, 2, stroke=1, fill=1)
    c.setFillColor(TEAL)
    c.setFont(BOLD, 5.8)
    c.drawCentredString(right - 20 * mm, top - 17.5 * mm, clip_text(status.upper() or "EM EDICAO", 36 * mm, BOLD, 5.8))
    c.setFillColor(MUTED)
    c.setFont(REGULAR, 5.4)
    c.drawRightString(right, top - 25 * mm, f"Pagina {page}")
    c.setStrokeColor(LINE)
    c.line(margin, top - 29 * mm, right, top - 29 * mm)
    return top - 34 * mm


def _light_party_box(c: Any, x: float, y_top: float, width: float, title: str, rows: list[tuple[str, str]]) -> None:
    height = 27 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(x, y_top - height, width, height, 2, stroke=1, fill=1)
    c.setFillColor(TEAL)
    c.rect(x, y_top - height, 2 * mm, height, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont(BOLD, 6.4)
    c.drawString(x + 5 * mm, y_top - 5 * mm, title.upper())
    line_y = y_top - 10 * mm
    for label, value in rows[:4]:
        c.setFillColor(MUTED)
        c.setFont(BOLD, 5.2)
        c.drawString(x + 5 * mm, line_y, f"{label.upper()}:")
        c.setFillColor(INK)
        c.setFont(REGULAR, 5.7)
        c.drawString(x + 23 * mm, line_y, clip_text(value or "-", width - 27 * mm, REGULAR, 5.7))
        line_y -= 4.2 * mm


def _light_section(c: Any, x: float, y_top: float, width: float, title: str) -> float:
    c.setFillColor(TEAL)
    c.rect(x, y_top - 6 * mm, 2 * mm, 6 * mm, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont(BOLD, 7.2)
    c.drawString(x + 5 * mm, y_top - 4.3 * mm, title.upper())
    c.setStrokeColor(LINE)
    c.line(x, y_top - 7 * mm, x + width, y_top - 7 * mm)
    return y_top - 10 * mm


def render_commercial_light(
    path: Path,
    logo: Path | None,
    *,
    kind: str,
    record: dict[str, Any],
    party: dict[str, Any],
    issuer: dict[str, Any],
) -> Path:
    size = A4
    width, _height = size
    margin = 12 * mm
    inner = width - 2 * margin
    lines = [dict(row) for row in list(record.get("linhas", []) or []) if isinstance(row, dict)]
    rows_per_page = 16
    pages = max(1, (len(lines) + rows_per_page - 1) // rows_per_page)
    is_rfq = kind == "Pedido de cotacao"
    c = pdf_canvas.Canvas(str(path), pagesize=size)

    party_title = "Fornecedor" if kind != "Orcamento" else "Cliente"
    party_name = _txt(party.get("nome") or party.get("empresa") or record.get("fornecedor"))
    document_date = _txt(record.get("data") or record.get("data_doc_ultima"))[:10] or date.today().isoformat()
    for page_index in range(pages):
        page_label = f"{page_index + 1}/{pages}"
        y = _light_header(c, size, logo, kind, _txt(record.get("numero")), _txt(record.get("estado")), page_label)
        gap = 4 * mm
        box_w = (inner - gap) / 2
        _light_party_box(
            c,
            margin,
            y,
            box_w,
            "Emitente",
            [
                ("Empresa", _txt(issuer.get("nome")) or "LUGEST"),
                ("NIF", _txt(issuer.get("nif"))),
                ("Morada", _txt(issuer.get("morada"))),
                ("Contacto", _txt(issuer.get("contacto"))),
            ],
        )
        _light_party_box(
            c,
            margin + box_w + gap,
            y,
            box_w,
            party_title,
            [
                ("Nome", party_name),
                ("NIF", _txt(party.get("nif"))),
                ("Morada", _txt(party.get("morada"))),
                ("Contacto", _txt(party.get("contacto") or party.get("email") or record.get("contacto"))),
            ],
        )
        y -= 32 * mm
        info = [
            ("Data", document_date),
            ("Entrega", _txt(record.get("data_entrega") or record.get("prazo_entrega_data") or record.get("prazo_entrega_texto")) or "A acordar"),
            ("Referencia", _txt(record.get("nota_cliente") or record.get("origem_cotacao")) or "-"),
        ]
        info_w = inner / len(info)
        for index, (label, value) in enumerate(info):
            c.setFillColor(MUTED)
            c.setFont(BOLD, 5.2)
            c.drawString(margin + index * info_w, y, label.upper())
            c.setFillColor(INK)
            c.setFont(REGULAR, 6.3)
            c.drawString(margin + index * info_w, y - 4 * mm, clip_text(value, info_w - 4 * mm, REGULAR, 6.3))
        y -= 10 * mm
        y = _light_section(c, margin, y, inner, "Artigos e condicoes" if not is_rfq else "Artigos para cotacao")

        if is_rfq:
            columns = [("Ref.", 24 * mm), ("Descricao", 82 * mm), ("Qtd.", 16 * mm), ("Un.", 12 * mm), ("Preco proposto", 27 * mm), ("Prazo", 25 * mm)]
        else:
            columns = [("Ref.", 23 * mm), ("Descricao", 72 * mm), ("Qtd.", 14 * mm), ("Un.", 10 * mm), ("Preco un.", 21 * mm), ("Desc.", 13 * mm), ("IVA", 11 * mm), ("Total", 22 * mm)]
        header_h = 8 * mm
        c.setFillColor(colors.HexColor("#E9EEF2"))
        c.rect(margin, y - header_h, inner, header_h, stroke=0, fill=1)
        c.setFillColor(TEAL)
        c.rect(margin, y - header_h, 2 * mm, header_h, stroke=0, fill=1)
        x = margin
        for label, column_w in columns:
            c.setFillColor(NAVY)
            c.setFont(BOLD, 5.3)
            c.drawString(x + 1.8 * mm, y - 5.2 * mm, clip_text(label.upper(), column_w - 3 * mm, BOLD, 5.3))
            x += column_w
        y -= header_h

        page_lines = lines[page_index * rows_per_page : (page_index + 1) * rows_per_page]
        for row_index, row in enumerate(page_lines):
            row_h = 9.5 * mm
            if row_index % 2:
                c.setFillColor(colors.HexColor("#F7F9FA"))
                c.rect(margin, y - row_h, inner, row_h, stroke=0, fill=1)
            c.setStrokeColor(LINE)
            c.line(margin, y - row_h, margin + inner, y - row_h)
            reference = _txt(row.get("ref") or row.get("ref_interna") or row.get("produto_codigo"))
            description = _txt(row.get("descricao"))
            quantity = _fmt(row.get("qtd"))
            unit = _txt(row.get("unid") or row.get("produto_unid")) or "UN"
            if is_rfq:
                cells = [reference, description, quantity, unit, "________________", "__________"]
            else:
                unit_price = row.get("preco", row.get("preco_unit", 0))
                discount = row.get("desconto", 0)
                vat = row.get("iva", record.get("iva_perc", 23))
                cells = [reference, description, quantity, unit, _eur(unit_price), f"{_fmt(discount)} %", f"{_fmt(vat)} %", _eur(row.get("total"))]
            x = margin
            for cell_index, ((_, column_w), value) in enumerate(zip(columns, cells)):
                font = BOLD if cell_index == 0 else REGULAR
                c.setFillColor(INK)
                c.setFont(font, 5.5)
                if cell_index == 1:
                    wrapped = wrap_text(value or "-", font, 5.5, column_w - 3.6 * mm, max_lines=2) or ["-"]
                    line_y = y - 3.8 * mm
                    for text_line in wrapped:
                        c.drawString(x + 1.8 * mm, line_y, text_line)
                        line_y -= 3 * mm
                else:
                    c.drawString(x + 1.8 * mm, y - 5.7 * mm, clip_text(value or "-", column_w - 3.6 * mm, font, 5.5))
                x += column_w
            y -= row_h

        if page_index == pages - 1:
            y -= 5 * mm
            if is_rfq:
                c.setFillColor(MUTED)
                c.setFont(REGULAR, 6)
                c.drawString(margin, y, "Validade da proposta: ____________________    Condicoes de pagamento: ______________________________")
                c.drawString(margin, y - 7 * mm, "Observacoes do fornecedor: ______________________________________________________________________")
            else:
                subtotal = float(record.get("subtotal", record.get("subtotal_bruto", 0)) or 0)
                total = float(record.get("total", 0) or 0)
                vat_value = max(total - subtotal, 0)
                totals_w = 66 * mm
                totals_x = margin + inner - totals_w
                c.setStrokeColor(LINE)
                c.setFillColor(WHITE)
                c.roundRect(totals_x, y - 24 * mm, totals_w, 24 * mm, 2, stroke=1, fill=1)
                for index, (label, value, bold) in enumerate(
                    [("Subtotal", _eur(subtotal), False), ("IVA", _eur(vat_value), False), ("Total", _eur(total), True)]
                ):
                    line_y = y - (6 + index * 7) * mm
                    c.setFillColor(NAVY if bold else MUTED)
                    c.setFont(BOLD if bold else REGULAR, 7 if bold else 6)
                    c.drawString(totals_x + 4 * mm, line_y, label.upper())
                    c.drawRightString(totals_x + totals_w - 4 * mm, line_y, value)

        _footer(c, size, "Documento de proposta | Layout economico de impressao", kind.upper(), page_label)
        c.showPage()

    c.save()
    return path


def render_route(path: Path, backend: LegacyBackend, logo: Path | None, order_number: str, guide_number: str) -> Path:
    size = landscape(A4)
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    y = _header(c, size, logo, "Folha de rota", "Execucao logistica e prova de entrega", "TR-EXEMPLO-001", "Planeada", page="1/1")
    margin = 12 * mm
    inner = width - 2 * margin
    gap = 3 * mm
    metrics = [("Paragens", "2", TEAL), ("Paletes", "4", GREEN), ("Peso bruto", "1 480 kg", AMBER), ("Volume", "5.8 m3", STEEL)]
    mw = (inner - 3 * gap) / 4
    for idx, (label, value, accent) in enumerate(metrics):
        _metric(c, margin + idx * (mw + gap), y, mw, label, value, accent)
    y -= 23 * mm
    y = _section(c, margin, y, inner, "Sequencia de distribuicao", "01")
    stops = [
        ("01", order_number or "LUGEST-EXEMPLO", "Cliente industrial Norte", "Zona Industrial | Porto", guide_number or "GT-EXEMPLO", "08:45", "PLANEADA"),
        ("02", "LUGEST-EXEMPLO-2", "Cliente industrial Centro", "Parque Empresarial | Leiria", "GT-EXEMPLO-2", "11:30", "PLANEADA"),
    ]
    for pos, enc, client, destination, guide, eta, state in stops:
        block_h = 34 * mm
        c.setFillColor(WHITE)
        c.setStrokeColor(LINE)
        c.roundRect(margin, y - block_h, inner, block_h, 3, stroke=1, fill=1)
        c.setFillColor(TEAL)
        c.roundRect(margin + 4 * mm, y - 13 * mm, 10 * mm, 10 * mm, 5 * mm, stroke=0, fill=1)
        c.setFillColor(WHITE)
        c.setFont(BOLD, 7)
        c.drawCentredString(margin + 9 * mm, y - 9.5 * mm, pos)
        c.setFillColor(INK)
        c.setFont(BOLD, 11)
        c.drawString(margin + 18 * mm, y - 8 * mm, client)
        c.setFont(REGULAR, 7)
        c.setFillColor(MUTED)
        c.drawString(margin + 18 * mm, y - 14 * mm, f"{enc} | {destination}")
        cols = [
            ("CHEGADA PREVISTA", eta),
            ("GUIA", guide),
            ("CARGA", "2 pal. | 740 kg"),
            ("ESTADO", state),
        ]
        cx = margin + 18 * mm
        cw = (inner - 24 * mm) / 4
        for label, value in cols:
            c.setFillColor(SURFACE)
            c.roundRect(cx, y - 29 * mm, cw - 2 * mm, 11 * mm, 2, stroke=0, fill=1)
            c.setFillColor(MUTED)
            c.setFont(BOLD, 4.8)
            c.drawString(cx + 2 * mm, y - 22 * mm, label)
            c.setFillColor(INK)
            c.setFont(BOLD, 6.7)
            c.drawString(cx + 2 * mm, y - 27 * mm, clip_text(value, cw - 6 * mm, BOLD, 6.7))
            cx += cw
        y -= block_h + 4 * mm
    y = _section(c, margin, y, inner, "Validacao da viagem", "02")
    _field(c, margin, y, inner * 0.46, "Motorista / assinatura", "")
    _field(c, margin + inner * 0.48, y, inner * 0.25, "Saida", "")
    _field(c, margin + inner * 0.75, y, inner * 0.25, "Regresso", "")
    _footer(c, size, "Documento operacional | Nao substitui a guia de transporte", "LOGISTICA V2", "1/1")
    c.save()
    return path


def render_product_sheet(path: Path, backend: LegacyBackend, logo: Path | None, product: dict[str, Any]) -> Path:
    size = A4
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    code = _txt(product.get("codigo")) or "PRD-EXEMPLO"
    description = _txt(product.get("descricao")) or "Produto industrial"
    y = _header(c, size, logo, code, description, "FICHA MESTRE DE PRODUTO", "Disponivel", page="1/1")
    margin = 12 * mm
    inner = width - 2 * margin
    gap = 3 * mm
    values = [
        ("Stock fisico", f"{_fmt(product.get('qty'))} {_txt(product.get('unid')) or 'UN'}", TEAL),
        ("Disponivel", f"{_fmt(product.get('disponivel', product.get('qty')))} {_txt(product.get('unid')) or 'UN'}", GREEN),
        ("Stock minimo", f"{_fmt(product.get('alerta'))} {_txt(product.get('unid')) or 'UN'}", AMBER),
        ("Localizacao", _txt(product.get("localizacao")) or "Por definir", STEEL),
    ]
    mw = (inner - 3 * gap) / 4
    for idx, spec in enumerate(values):
        _metric(c, margin + idx * (mw + gap), y, mw, *spec)
    y -= 23 * mm
    y = _section(c, margin, y, inner, "Identificacao tecnica", "01")
    half = (inner - gap) / 2
    fields = [
        ("Categoria / subcategoria", f"{_txt(product.get('categoria')) or '-'} | {_txt(product.get('subcategoria')) or '-'}"),
        ("Tipo / unidade", f"{_txt(product.get('tipo')) or '-'} | {_txt(product.get('unid')) or 'UN'}"),
        ("Fabricante / modelo", f"{_txt(product.get('fabricante')) or '-'} | {_txt(product.get('modelo')) or '-'}"),
        ("Dimensoes / peso", f"{_txt(product.get('dimensoes')) or '-'} | {_fmt(product.get('peso_unid'))} kg"),
    ]
    for idx, (label, value) in enumerate(fields):
        col, row = idx % 2, idx // 2
        _field(c, margin + col * (half + gap), y - row * 20 * mm, half, label, value)
    y -= 42 * mm
    y = _section(c, margin, y, inner, "Rastreabilidade e abastecimento", "02")
    _field(c, margin, y, half, "Referencia fornecedor", _txt(product.get("ref_fornecedor")) or "-")
    _field(c, margin + half + gap, y, half, "Fornecedor preferencial", _txt(product.get("fornecedor")) or "-")
    y -= 20 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(margin, y - 22 * mm, inner, 22 * mm, 3, stroke=1, fill=1)
    _barcode(c, code, margin + 4 * mm, y - 19 * mm, inner - 8 * mm, 16 * mm)
    y -= 28 * mm
    y = _section(c, margin, y, inner, "Movimentos recentes", "03")
    movements = list(product.get("movimentos", []) or [])[:7]
    headers = [("DATA", 28 * mm), ("TIPO", 34 * mm), ("OPERADOR", 34 * mm), ("QTD", 18 * mm), ("ANTES", 18 * mm), ("DEPOIS", 18 * mm), ("ORIGEM", inner - 150 * mm)]
    c.setFillColor(colors.HexColor("#E9EEF2"))
    c.rect(margin, y - 8 * mm, inner, 8 * mm, stroke=0, fill=1)
    c.setFillColor(TEAL)
    c.rect(margin, y - 8 * mm, 2 * mm, 8 * mm, stroke=0, fill=1)
    x = margin
    c.setFillColor(NAVY)
    c.setFont(BOLD, 5.5)
    for label, col_w in headers:
        c.drawString(x + 2 * mm, y - 5.2 * mm, label)
        x += col_w
    y -= 8 * mm
    for idx, row in enumerate(movements or [{"data": "-", "tipo": "Sem movimentos"}]):
        rh = 9 * mm
        c.setFillColor(WHITE if idx % 2 == 0 else SURFACE)
        c.setStrokeColor(LINE)
        c.rect(margin, y - rh, inner, rh, stroke=1, fill=1)
        vals = [row.get("data", "-"), row.get("tipo", "-"), row.get("operador", "-"), row.get("qtd", "-"), row.get("antes", "-"), row.get("depois", "-"), row.get("origem", "-")]
        x = margin
        c.setFillColor(INK)
        c.setFont(REGULAR, 5.7)
        for value, (_label, col_w) in zip(vals, headers):
            c.drawString(x + 2 * mm, y - 5.8 * mm, clip_text(_txt(value), col_w - 4 * mm, REGULAR, 5.7))
            x += col_w
        y -= rh
    _footer(c, size, "Dados sincronizados com stock e compras", "PRODUTO V2", "1/1")
    c.save()
    return path


def render_compact_label(path: Path, logo: Path | None, *, title: str, code: str, description: str, meta: list[tuple[str, str]], barcode_value: str) -> Path:
    size = (150 * mm, 100 * mm)
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    margin = 7 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(margin, height - 29 * mm, width - 2 * margin, 22 * mm, 3, stroke=1, fill=1)
    c.setFillColor(TEAL)
    c.rect(margin, height - 9 * mm, width - 2 * margin, 2 * mm, stroke=0, fill=1)
    c.rect(margin, height - 29 * mm, 2 * mm, 20 * mm, stroke=0, fill=1)
    _draw_logo(c, logo, margin + 4 * mm, height - 25 * mm, 28 * mm, 14 * mm)
    c.setFillColor(NAVY)
    c.setFont(BOLD, 12)
    c.drawString(margin + 37 * mm, height - 16 * mm, clip_text(title, width - margin - (margin + 39 * mm), BOLD, 12))
    c.setFillColor(MUTED)
    c.setFont(REGULAR, 6)
    c.drawString(margin + 37 * mm, height - 22 * mm, clip_text(description, width - margin - (margin + 39 * mm), REGULAR, 6))
    y = height - 34 * mm
    c.setFillColor(INK)
    c.setFont(BOLD, 17)
    c.drawString(margin, y, clip_text(code, width - 2 * margin, BOLD, 17))
    y -= 5 * mm
    gap = 2 * mm
    cols = 2
    fw = (width - 2 * margin - gap) / cols
    for idx, (label, value) in enumerate(meta[:4]):
        col, row = idx % cols, idx // cols
        _field(c, margin + col * (fw + gap), y - row * 17 * mm, fw, label, value, height=14 * mm)
    bar_y = 8 * mm
    c.setFillColor(WHITE)
    c.setStrokeColor(LINE)
    c.roundRect(margin, bar_y, width - 2 * margin, 21 * mm, 3, stroke=1, fill=1)
    _barcode(c, barcode_value, margin + 4 * mm, bar_y + 2 * mm, width - 2 * margin - 8 * mm, 16 * mm)
    c.save()
    return path


def render_pallet_label(path: Path, logo: Path | None, row: dict[str, Any]) -> Path:
    size = landscape(A4)
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    y = _header(c, size, logo, "Etiqueta de palete", "Fluxo interno de producao", f"PLT-{_txt(row.get('encomenda'))}", "Em transferencia", page="1/1")
    margin = 12 * mm
    inner = width - 2 * margin
    _metric(c, margin, y, inner * 0.32, "Destino", _txt(row.get("proximo_posto")) or "Expedicao", TEAL, 22 * mm)
    _metric(c, margin + inner * 0.34, y, inner * 0.2, "Referencias", "1", GREEN, 22 * mm)
    _metric(c, margin + inner * 0.56, y, inner * 0.2, "Quantidade", _fmt(row.get("quantidade")), AMBER, 22 * mm)
    _metric(c, margin + inner * 0.78, y, inner * 0.22, "Origem", _txt(row.get("posto_origem")) or "Producao", STEEL, 22 * mm)
    y -= 28 * mm
    y = _section(c, margin, y, inner, "Conteudo da palete", "01")
    headers = [("REF. INTERNA", 44 * mm), ("REF. EXTERNA", 60 * mm), ("DESCRICAO", 74 * mm), ("MATERIAL", 30 * mm), ("ESP.", 20 * mm), ("QTD", 20 * mm), ("OPP", inner - 248 * mm)]
    c.setFillColor(colors.HexColor("#E9EEF2"))
    c.rect(margin, y - 9 * mm, inner, 9 * mm, stroke=0, fill=1)
    c.setFillColor(TEAL)
    c.rect(margin, y - 9 * mm, 2 * mm, 9 * mm, stroke=0, fill=1)
    x = margin
    c.setFillColor(NAVY)
    c.setFont(BOLD, 6)
    for label, col_w in headers:
        c.drawString(x + 2 * mm, y - 5.8 * mm, label)
        x += col_w
    y -= 9 * mm
    c.setFillColor(SURFACE)
    c.setStrokeColor(LINE)
    c.rect(margin, y - 18 * mm, inner, 18 * mm, stroke=1, fill=1)
    vals = [row.get("ref_interna"), row.get("ref_externa"), row.get("descricao"), row.get("material"), row.get("espessura"), row.get("quantidade"), row.get("opp")]
    x = margin
    c.setFillColor(INK)
    c.setFont(BOLD, 7.3)
    for value, (_label, col_w) in zip(vals, headers):
        c.drawString(x + 2 * mm, y - 10.5 * mm, clip_text(_txt(value) or "-", col_w - 4 * mm, BOLD, 7.3))
        x += col_w
    y -= 25 * mm
    y = _section(c, margin, y, inner, "Confirmacao", "02")
    _field(c, margin, y, inner * 0.31, "Preparado por", "")
    _field(c, margin + inner * 0.33, y, inner * 0.31, "Recebido por", "")
    _field(c, margin + inner * 0.66, y, inner * 0.34, "Data / hora", "")
    _footer(c, size, "Etiqueta operacional controlada", "PALETE V2", "1/1")
    c.save()
    return path


def render_history(path: Path, logo: Path | None, rows: list[dict[str, Any]]) -> Path:
    size = landscape(A4)
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    margin = 12 * mm
    inner = width - 2 * margin
    per_page = 11
    chunks = [rows[i : i + per_page] for i in range(0, len(rows), per_page)] or [[]]
    for page_no, chunk in enumerate(chunks, 1):
        if page_no > 1:
            c.showPage()
        y = _header(c, size, logo, "Historico de materia-prima", "Movimentos, lotes e rastreabilidade", "STOCK / AUDITORIA", "Atualizado", page=f"{page_no}/{len(chunks)}")
        metrics = [("Movimentos", str(len(rows)), TEAL), ("Entradas", str(sum(1 for r in rows if "ENTR" in _txt(r.get("acao")).upper() or "ADIC" in _txt(r.get("acao")).upper())), GREEN), ("Baixas", str(sum(1 for r in rows if "BAIXA" in _txt(r.get("acao")).upper())), AMBER)]
        mw = (inner - 6 * mm) / 3
        for idx, spec in enumerate(metrics):
            _metric(c, margin + idx * (mw + 3 * mm), y, mw, *spec, height=15 * mm)
        y -= 20 * mm
        y = _section(c, margin, y, inner, "Registo cronologico", "01")
        headers = [("DATA", 32 * mm), ("ACAO", 24 * mm), ("OPERADOR", 30 * mm), ("MATERIAL", 36 * mm), ("ESP.", 14 * mm), ("DIM.", 38 * mm), ("LOTE", 32 * mm), ("QTD", 14 * mm), ("DETALHE", inner - 220 * mm)]
        c.setFillColor(colors.HexColor("#E9EEF2"))
        c.rect(margin, y - 8 * mm, inner, 8 * mm, stroke=0, fill=1)
        c.setFillColor(TEAL)
        c.rect(margin, y - 8 * mm, 2 * mm, 8 * mm, stroke=0, fill=1)
        x = margin
        c.setFillColor(NAVY)
        c.setFont(BOLD, 5.5)
        for label, col_w in headers:
            c.drawString(x + 2 * mm, y - 5.1 * mm, label)
            x += col_w
        y -= 8 * mm
        for idx, row in enumerate(chunk):
            rh = 10 * mm
            c.setFillColor(WHITE if idx % 2 == 0 else SURFACE)
            c.setStrokeColor(LINE)
            c.rect(margin, y - rh, inner, rh, stroke=1, fill=1)
            vals = [row.get("data"), row.get("acao"), row.get("operador"), row.get("material"), row.get("espessura"), row.get("dimensao"), row.get("lote"), row.get("qtd"), row.get("detalhes")]
            x = margin
            c.setFillColor(INK)
            c.setFont(REGULAR, 5.6)
            for value, (_label, col_w) in zip(vals, headers):
                c.drawString(x + 2 * mm, y - 6.2 * mm, clip_text(_txt(value) or "-", col_w - 4 * mm, REGULAR, 5.6))
                x += col_w
            y -= rh
        _footer(c, size, "Historico consolidado para rastreabilidade", "MATERIA-PRIMA V2", f"{page_no}/{len(chunks)}")
    c.save()
    return path


def render_separation(path: Path, logo: Path | None, rows: list[dict[str, Any]]) -> Path:
    size = landscape(A4)
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    y = _header(c, size, logo, "Separacao de materia-prima", "Preparacao por posto e horizonte de 5 dias", "ASSISTENTE DE MATERIAL", "Plano atual", page="1/1")
    margin = 12 * mm
    inner = width - 2 * margin
    metrics = [("Linhas", str(len(rows)), TEAL), ("Postos", str(len({_txt(r.get('posto_trabalho')) for r in rows if _txt(r.get('posto_trabalho'))})), GREEN), ("Confirmadas", str(sum(1 for r in rows if r.get("visto_sep_checked"))), AMBER), ("Pendentes", str(sum(1 for r in rows if not r.get("visto_sep_checked"))), STEEL)]
    mw = (inner - 9 * mm) / 4
    for idx, spec in enumerate(metrics):
        _metric(c, margin + idx * (mw + 3 * mm), y, mw, *spec)
    y -= 24 * mm
    y = _section(c, margin, y, inner, "Plano de separacao", "01")
    if not rows:
        c.setFillColor(SURFACE)
        c.setStrokeColor(LINE)
        c.roundRect(margin, y - 70 * mm, inner, 70 * mm, 4, stroke=1, fill=1)
        c.setFillColor(TEAL)
        c.circle(width / 2, y - 25 * mm, 8 * mm, stroke=0, fill=1)
        c.setFillColor(WHITE)
        c.setFont(BOLD, 13)
        c.drawCentredString(width / 2, y - 29 * mm, "OK")
        c.setFillColor(INK)
        c.setFont(BOLD, 15)
        c.drawCentredString(width / 2, y - 45 * mm, "Sem necessidades de separacao no horizonte atual")
        c.setFillColor(MUTED)
        c.setFont(REGULAR, 8)
        c.drawCentredString(width / 2, y - 53 * mm, "O documento fica numa unica pagina e identifica explicitamente o estado vazio.")
    _footer(c, size, "Horizonte calculado a partir do planeamento", "SEPARACAO V2", "1/1")
    c.save()
    return path


def render_planning(path: Path, logo: Path | None, rows: list[dict[str, Any]]) -> Path:
    size = landscape(A4)
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    y = _header(c, size, logo, "Prazos do fluxo produtivo", "Conclusao prevista por encomenda", "PLANEAMENTO / CORTE LASER", "Sincronizado", page="1/1")
    margin = 12 * mm
    inner = width - 2 * margin
    metrics = [("Encomendas", str(len(rows)), TEAL), ("Fechadas", str(sum(1 for r in rows if "fech" in _txt(r.get("estado")).lower())), GREEN), ("Parciais", str(sum(1 for r in rows if "parcial" in _txt(r.get("estado")).lower())), AMBER), ("Por planear", str(sum(1 for r in rows if "planear" in _txt(r.get("estado")).lower())), RED)]
    mw = (inner - 9 * mm) / 4
    for idx, spec in enumerate(metrics):
        _metric(c, margin + idx * (mw + 3 * mm), y, mw, *spec)
    y -= 24 * mm
    y = _section(c, margin, y, inner, "Prioridades e datas previstas", "01")
    headers = [("ENCOMENDA", 40 * mm), ("CLIENTE", 52 * mm), ("ENTREGA", 28 * mm), ("GRUPOS", 25 * mm), ("PLANEADO", 42 * mm), ("FIM FLUXO", 42 * mm), ("ESTADO", inner - 229 * mm)]
    c.setFillColor(colors.HexColor("#E9EEF2"))
    c.rect(margin, y - 9 * mm, inner, 9 * mm, stroke=0, fill=1)
    c.setFillColor(TEAL)
    c.rect(margin, y - 9 * mm, 2 * mm, 9 * mm, stroke=0, fill=1)
    x = margin
    c.setFillColor(NAVY)
    c.setFont(BOLD, 5.8)
    for label, col_w in headers:
        c.drawString(x + 2 * mm, y - 5.8 * mm, label)
        x += col_w
    y -= 9 * mm
    for idx, row in enumerate(rows[:12]):
        rh = 11 * mm
        c.setFillColor(WHITE if idx % 2 == 0 else SURFACE)
        c.setStrokeColor(LINE)
        c.rect(margin, y - rh, inner, rh, stroke=1, fill=1)
        vals = [row.get("numero"), row.get("cliente"), row.get("data_entrega"), row.get("grupos_txt"), row.get("planeado_txt"), row.get("fim_txt"), row.get("estado")]
        x = margin
        c.setFillColor(INK)
        c.setFont(BOLD if idx < 3 else REGULAR, 6.2)
        for value, (_label, col_w) in zip(vals, headers):
            c.drawString(x + 2 * mm, y - 6.8 * mm, clip_text(_txt(value) or "-", col_w - 4 * mm, BOLD if idx < 3 else REGULAR, 6.2))
            x += col_w
        y -= rh
    _footer(c, size, "Prioridade calculada pelo planeamento e estado operacional", "PLANEAMENTO V2", "1/1")
    c.save()
    return path


def render_quality_dossier(path: Path, backend: LegacyBackend, logo: Path | None) -> Path:
    size = A4
    width, height = size
    c = pdf_canvas.Canvas(str(path), pagesize=size)
    summary = dict(backend.quality_summary() or {})
    checklist = list(backend.quality_iso_checklist() or [])
    ncs = list(backend.quality_nc_rows("", "Todos") or [])
    docs = list(backend.quality_document_rows("") or [])
    pages = 3
    y = _header(c, size, logo, "Dossier da Qualidade", "Sistema de gestao e melhoria continua", "ISO 9001 / REV. 00", "Controlado", page=f"1/{pages}")
    margin = 12 * mm
    inner = width - 2 * margin
    metrics = [
        ("NC abertas", str(summary.get("nc_abertas", summary.get("nc_open", len(ncs)))), RED),
        ("Documentos", str(len(docs)), TEAL),
        ("Checklist", str(len(checklist)), GREEN),
        ("Auditoria", str(summary.get("audit_events", len(backend.audit_rows("", limit=3000)))), STEEL),
    ]
    mw = (inner - 9 * mm) / 4
    for idx, spec in enumerate(metrics):
        _metric(c, margin + idx * (mw + 3 * mm), y, mw, *spec)
    y -= 24 * mm
    y = _section(c, margin, y, inner, "Estado do sistema", "01")
    c.setFillColor(SURFACE)
    c.setStrokeColor(LINE)
    c.roundRect(margin, y - 52 * mm, inner, 52 * mm, 4, stroke=1, fill=1)
    c.setFillColor(INK)
    c.setFont(BOLD, 16)
    c.drawString(margin + 8 * mm, y - 14 * mm, "Qualidade integrada no fluxo industrial")
    c.setFillColor(MUTED)
    c.setFont(REGULAR, 8)
    lines = [
        "Rececao, nao conformidades, fornecedores, documentos e auditoria num unico dossier.",
        "Os indicadores seguintes refletem o estado atual da base de dados.",
        "Documento preparado para revisao de gestao e acompanhamento operacional.",
    ]
    line_y = y - 23 * mm
    for line in lines:
        c.drawString(margin + 8 * mm, line_y, line)
        line_y -= 6 * mm
    y -= 60 * mm
    y = _section(c, margin, y, inner, "Pontos de controlo", "02")
    for idx, item in enumerate(checklist[:6]):
        col, row = idx % 2, idx // 2
        x = margin + col * ((inner - 3 * mm) / 2 + 3 * mm)
        top = y - row * 18 * mm
        w = (inner - 3 * mm) / 2
        c.setFillColor(WHITE)
        c.setStrokeColor(LINE)
        c.roundRect(x, top - 14 * mm, w, 14 * mm, 2, stroke=1, fill=1)
        c.setFillColor(GREEN if "ok" in _txt(item.get("estado")).lower() else AMBER)
        c.circle(x + 6 * mm, top - 7 * mm, 3 * mm, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont(BOLD, 6.4)
        c.drawString(x + 12 * mm, top - 5.2 * mm, clip_text(_txt(item.get("area")) or "Controlo", w - 15 * mm, BOLD, 6.4))
        c.setFillColor(MUTED)
        c.setFont(REGULAR, 5.2)
        c.drawString(x + 12 * mm, top - 10 * mm, clip_text(_txt(item.get("evidencia")) or "-", w - 15 * mm, REGULAR, 5.2))
    _footer(c, size, "Dossier para revisao do sistema", "QUALIDADE V2", f"1/{pages}")

    c.showPage()
    y = _header(c, size, logo, "Nao conformidades", "Estado, responsabilidade e decisao", "DOSSIER DA QUALIDADE", "Acompanhamento", page=f"2/{pages}")
    y = _section(c, margin, y, inner, "Registo de nao conformidades", "03")
    for idx, row in enumerate(ncs[:8] or [{"id": "-", "descricao": "Sem nao conformidades registadas", "estado": "OK"}]):
        bh = 23 * mm
        c.setFillColor(WHITE if idx % 2 == 0 else SURFACE)
        c.setStrokeColor(LINE)
        c.roundRect(margin, y - bh, inner, bh, 2, stroke=1, fill=1)
        c.setFillColor(RED if "abert" in _txt(row.get("estado")).lower() else GREEN)
        c.rect(margin, y - bh, 3 * mm, bh, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont(BOLD, 8)
        c.drawString(margin + 6 * mm, y - 7 * mm, f"{_txt(row.get('id')) or '-'} | {_txt(row.get('estado')) or '-'}")
        c.setFont(REGULAR, 6.2)
        desc_lines = wrap_text(_txt(row.get("descricao")) or "-", REGULAR, 6.2, inner - 12 * mm, max_lines=2)
        ly = y - 13 * mm
        for line in desc_lines:
            c.drawString(margin + 6 * mm, ly, line)
            ly -= 3.6 * mm
        y -= bh + 3 * mm
    _footer(c, size, "Nao conformidades e acoes corretivas", "QUALIDADE V2", f"2/{pages}")

    c.showPage()
    y = _header(c, size, logo, "Documentos e auditoria", "Evidencia, validade e rastreabilidade", "DOSSIER DA QUALIDADE", "Controlado", page=f"3/{pages}")
    y = _section(c, margin, y, inner, "Documentacao do sistema", "04")
    for idx, row in enumerate(docs[:10] or [{"id": "-", "tipo": "-", "titulo": "Sem documentos registados"}]):
        rh = 13 * mm
        c.setFillColor(WHITE if idx % 2 == 0 else SURFACE)
        c.setStrokeColor(LINE)
        c.rect(margin, y - rh, inner, rh, stroke=1, fill=1)
        c.setFillColor(INK)
        c.setFont(BOLD, 6.4)
        c.drawString(margin + 3 * mm, y - 5 * mm, f"{_txt(row.get('id')) or '-'} | {_txt(row.get('tipo')) or '-'}")
        c.setFont(REGULAR, 6)
        c.drawString(margin + 3 * mm, y - 10 * mm, clip_text(_txt(row.get("titulo")) or "-", inner - 6 * mm, REGULAR, 6))
        y -= rh
    _footer(c, size, "Documentos, versoes e evidencia de auditoria", "QUALIDADE V2", f"3/{pages}")
    c.save()
    return path


def _stock_state(record: dict[str, Any], available: float, minimum: float = 0) -> tuple[str, Any]:
    if bool(record.get("quality_blocked")):
        return "BLOQUEADO", RED
    if available <= 0:
        return "SEM STOCK", RED
    if minimum > 0 and available <= minimum:
        return "REPOR", AMBER
    quality = _txt(record.get("quality_status")).upper()
    if quality and quality not in {"APROVADO", "OK", "CONFORME", "LIBERTADO"}:
        return quality.replace("_", " "), AMBER
    return "DISPONIVEL", GREEN


def _material_dimension(record: dict[str, Any]) -> str:
    values = []
    for key in ("comprimento", "largura", "altura"):
        value = float(record.get(key, 0) or 0)
        if value > 0:
            values.append(_fmt(value))
    if values:
        return " x ".join(values) + " mm"
    diameter = float(record.get("diametro", 0) or 0)
    metres = float(record.get("metros", 0) or 0)
    if diameter > 0:
        return f"Ø {_fmt(diameter)} mm"
    if metres > 0:
        return f"{_fmt(metres)} m"
    return "-"


def _render_stock_pages(
    path: Path,
    logo: Path | None,
    *,
    title: str,
    subtitle: str,
    document: str,
    section_title: str,
    metrics: list[tuple[str, str, Any]],
    columns: list[tuple[str, float]],
    rows: list[tuple[list[str], tuple[str, Any]]],
) -> Path:
    size = landscape(A4)
    width, _height = size
    margin = 12 * mm
    inner = width - 2 * margin
    rows_per_page = 12
    pages = max(1, (len(rows) + rows_per_page - 1) // rows_per_page)
    c = pdf_canvas.Canvas(str(path), pagesize=size)

    for page_index in range(pages):
        page_label = f"{page_index + 1}/{pages}"
        y = _header(c, size, logo, title, subtitle, document, "Mapa controlado", page=page_label)
        gap = 3 * mm
        metric_w = (inner - gap * (len(metrics) - 1)) / len(metrics)
        for index, (label, value, accent) in enumerate(metrics):
            _metric(c, margin + index * (metric_w + gap), y, metric_w, label, value, accent, height=15 * mm)
        y -= 20 * mm
        y = _section(c, margin, y, inner, section_title, "01")

        header_h = 8 * mm
        c.setFillColor(colors.HexColor("#E9EEF2"))
        c.rect(margin, y - header_h, inner, header_h, stroke=0, fill=1)
        c.setFillColor(TEAL)
        c.rect(margin, y - header_h, 2 * mm, header_h, stroke=0, fill=1)
        x = margin
        for label, column_w in columns:
            c.setFillColor(NAVY)
            c.setFont(BOLD, 5.3)
            c.drawString(x + 1.8 * mm, y - 5.2 * mm, clip_text(label.upper(), column_w - 3.6 * mm, BOLD, 5.3))
            x += column_w
        y -= header_h

        page_rows = rows[page_index * rows_per_page : (page_index + 1) * rows_per_page]
        for row_index, (cells, state) in enumerate(page_rows):
            row_h = 8.8 * mm
            c.setFillColor(WHITE if row_index % 2 == 0 else SURFACE)
            c.setStrokeColor(LINE)
            c.rect(margin, y - row_h, inner, row_h, stroke=1, fill=1)
            x = margin
            for cell_index, ((_, column_w), value) in enumerate(zip(columns, cells)):
                c.setFillColor(state[1] if cell_index == len(cells) - 1 else INK)
                c.setFont(BOLD if cell_index in {0, len(cells) - 1} else REGULAR, 5.7)
                font = BOLD if cell_index in {0, len(cells) - 1} else REGULAR
                c.drawString(x + 1.8 * mm, y - 5.7 * mm, clip_text(value or "-", column_w - 3.6 * mm, font, 5.7))
                x += column_w
            y -= row_h

        _footer(c, size, "Inventario operacional | Valores sujeitos aos movimentos de stock", document, page_label)
        c.showPage()

    c.save()
    return path


def render_material_stock(path: Path, logo: Path | None, source_rows: list[dict[str, Any]]) -> Path:
    records = [dict(item.get("record") or {}) for item in source_rows if isinstance(item, dict)]
    total_available = sum(max(float(row.get("quantidade", 0) or 0) - float(row.get("reservado", 0) or 0), 0) for row in records)
    total_reserved = sum(float(row.get("reservado", 0) or 0) for row in records)
    blocked = sum(1 for row in records if bool(row.get("quality_blocked")))
    stock_value = sum(float(row.get("quantidade", 0) or 0) * float(row.get("preco_unid", 0) or 0) for row in records)
    columns = [
        ("ID", 22 * mm),
        ("Formato", 20 * mm),
        ("Material", 34 * mm),
        ("Dimensao", 45 * mm),
        ("Esp.", 14 * mm),
        ("Qtd.", 14 * mm),
        ("Reserv.", 16 * mm),
        ("Dispon.", 16 * mm),
        ("Lote interno", 29 * mm),
        ("Local", 24 * mm),
        ("Estado", 38.8 * mm),
    ]
    table_rows = []
    for row in records:
        quantity = float(row.get("quantidade", 0) or 0)
        reserved = float(row.get("reservado", 0) or 0)
        available = max(quantity - reserved, 0)
        state = _stock_state(row, available)
        table_rows.append(
            (
                [
                    _txt(row.get("id")),
                    _txt(row.get("formato")),
                    _txt(row.get("material")),
                    _material_dimension(row),
                    _txt(row.get("espessura")) or "-",
                    _fmt(quantity),
                    _fmt(reserved),
                    _fmt(available),
                    _txt(row.get("lote_interno")),
                    _txt(row.get("Localizacao")),
                    state[0],
                ],
                state,
            )
        )
    table_rows.sort(key=lambda item: (item[1][0] == "DISPONIVEL", item[0][2], item[0][0]))
    return _render_stock_pages(
        path,
        logo,
        title="Stock de materia-prima",
        subtitle="Disponibilidade, lotes, geometria e controlo de qualidade",
        document="INVENTARIO MP V2",
        section_title="Mapa detalhado de materia-prima",
        metrics=[
            ("Referencias", str(len(records)), TEAL),
            ("Disponivel", _fmt(total_available), GREEN),
            ("Reservado", _fmt(total_reserved), AMBER),
            ("Bloqueado", str(blocked), RED if blocked else GREEN),
            ("Valor de stock", f"{stock_value:,.2f} EUR", STEEL),
        ],
        columns=columns,
        rows=table_rows,
    )


def render_product_stock(path: Path, logo: Path | None, source_rows: list[dict[str, Any]]) -> Path:
    records = [dict(row) for row in source_rows if isinstance(row, dict)]
    available = sum(float(row.get("available_qty", row.get("qty", 0)) or 0) for row in records)
    low_stock = sum(
        1
        for row in records
        if float(row.get("alerta", 0) or 0) > 0
        and float(row.get("available_qty", row.get("qty", 0)) or 0) <= float(row.get("alerta", 0) or 0)
    )
    blocked = sum(1 for row in records if bool(row.get("quality_blocked")))
    stock_value = sum(float(row.get("valor_stock", 0) or 0) for row in records)
    columns = [
        ("Codigo", 25 * mm),
        ("Descricao", 60 * mm),
        ("Categoria", 30 * mm),
        ("Tipo", 35 * mm),
        ("Un.", 12 * mm),
        ("Stock", 16 * mm),
        ("Min.", 15 * mm),
        ("Dispon.", 17 * mm),
        ("Fabricante / modelo", 32 * mm),
        ("Estado", 31.8 * mm),
    ]
    table_rows = []
    for row in records:
        stock = float(row.get("qty", 0) or 0)
        minimum = float(row.get("alerta", 0) or 0)
        item_available = float(row.get("available_qty", stock) or 0)
        state = _stock_state(row, item_available, minimum)
        maker = " / ".join(part for part in (_txt(row.get("fabricante")), _txt(row.get("modelo"))) if part)
        table_rows.append(
            (
                [
                    _txt(row.get("codigo")),
                    _txt(row.get("descricao")),
                    _txt(row.get("categoria")),
                    _txt(row.get("type_display")),
                    _txt(row.get("unid")) or "UN",
                    _fmt(stock),
                    _fmt(minimum),
                    _fmt(item_available),
                    maker,
                    state[0],
                ],
                state,
            )
        )
    table_rows.sort(key=lambda item: (item[1][0] == "DISPONIVEL", item[0][2], item[0][0]))
    return _render_stock_pages(
        path,
        logo,
        title="Stock de produtos",
        subtitle="Disponibilidade, familias, niveis minimos e estado operacional",
        document="INVENTARIO PRODUTO V2",
        section_title="Mapa detalhado de produtos",
        metrics=[
            ("Referencias", str(len(records)), TEAL),
            ("Disponivel", _fmt(available), GREEN),
            ("A repor", str(low_stock), AMBER if low_stock else GREEN),
            ("Bloqueado", str(blocked), RED if blocked else GREEN),
            ("Valor de stock", f"{stock_value:,.2f} EUR", STEEL),
        ],
        columns=columns,
        rows=table_rows,
    )


def _output_dir() -> Path:
    desktop = Path.home() / "Desktop"
    base = desktop / f"Propostas PDFs LUGEST LIGHT - {date.today().isoformat()}"
    if not base.exists():
        return base
    return desktop / f"{base.name} - {datetime.now().strftime('%H%M%S')}"


def main() -> int:
    backend = LegacyBackend()
    backend._replace_data_cache(copy.deepcopy(backend.ensure_data()))
    backend._save = lambda *args, **kwargs: None
    branding = dict(backend.branding_settings() or {})
    logo_raw = _txt(branding.get("logo_path"))
    logo = Path(logo_raw) if logo_raw and Path(logo_raw).exists() else None
    out = _output_dir()
    out.mkdir(parents=True, exist_ok=False)

    data = backend.ensure_data()
    products = [row for row in data.get("produtos", []) if isinstance(row, dict)]
    materials = [row for row in data.get("materiais", []) if isinstance(row, dict)]
    orders = [row for row in data.get("encomendas", []) if isinstance(row, dict)]
    product = backend.product_detail(_txt(products[0].get("codigo"))) if products else {}
    material = materials[0] if materials else {}
    order = max(orders, key=lambda row: len(list(backend.desktop_main.encomenda_pecas(row) or [])), default={})
    pieces = list(backend.desktop_main.encomenda_pecas(order) or [])
    piece = next((row for row in pieces if isinstance(row, dict)), {})
    label_row = backend._operator_label_row(order, piece, source_posto="Corte Laser") if order and piece else {}
    order_number = _txt(order.get("numero"))
    guide_number = _txt((data.get("expedicoes") or [{}])[0].get("numero"))
    material_stock_rows = list(backend.material_rows() or [])
    product_stock_rows = list(backend.product_rows() or [])
    purchase_notes = [row for row in data.get("notas_encomenda", []) if isinstance(row, dict)]
    purchase_note = max(purchase_notes, key=lambda row: len(list(row.get("linhas", []) or [])), default={})
    supplier_id = _txt(purchase_note.get("fornecedor_id") or purchase_note.get("fornecedor"))
    supplier = next(
        (dict(row) for row in data.get("fornecedores", []) if isinstance(row, dict) and _txt(row.get("id")) == supplier_id),
        {},
    )
    quotes = [row for row in data.get("orcamentos", []) if isinstance(row, dict)]
    quote = max(quotes, key=lambda row: len(list(row.get("linhas", []) or [])), default={})
    issuer = dict(branding.get("guia_emitente") or {})

    outputs = [
        ("01_Folha_Rota_Transporte_Light_V2.pdf", lambda p: render_route(p, backend, logo, order_number, guide_number)),
        ("02_Ficha_Produto_Light_V2.pdf", lambda p: render_product_sheet(p, backend, logo, product)),
        (
            "03_Etiqueta_Produto_Light_V2.pdf",
            lambda p: render_compact_label(
                p,
                logo,
                title="Etiqueta de produto",
                code=_txt(product.get("codigo")) or "PRD-EXEMPLO",
                description=_txt(product.get("descricao")) or "Produto industrial",
                meta=[
                    ("Stock", f"{_fmt(product.get('qty'))} {_txt(product.get('unid')) or 'UN'}"),
                    ("Localizacao", _txt(product.get("localizacao")) or "-"),
                    ("Categoria", _txt(product.get("categoria")) or "-"),
                    ("Estado", "Disponivel"),
                ],
                barcode_value=_txt(product.get("codigo")) or "PRD-EXEMPLO",
            ),
        ),
        (
            "04_Etiqueta_Materia_Prima_Light_V2.pdf",
            lambda p: render_compact_label(
                p,
                logo,
                title="Etiqueta de materia-prima",
                code=_txt(material.get("id")) or "MAT-EXEMPLO",
                description=f"{_txt(material.get('material')) or '-'} | {_txt(material.get('espessura')) or '-'} mm",
                meta=[
                    ("Formato", _txt(material.get("formato")) or "-"),
                    (
                        "Dimensao",
                        f"{_fmt(material.get('comprimento'))} x {_fmt(material.get('largura'))} x {_txt(material.get('espessura')) or '-'} mm",
                    ),
                    ("Lote", _txt(material.get("lote_interno")) or "-"),
                    ("Quantidade", _fmt(material.get("quantidade"))),
                ],
                barcode_value=_txt(material.get("id")) or "MAT-EXEMPLO",
            ),
        ),
        (
            "05_Etiqueta_Operador_Light_V2.pdf",
            lambda p: render_compact_label(
                p,
                logo,
                title="Etiqueta OPP",
                code=_txt(label_row.get("ref_interna")) or "REF-EXEMPLO",
                description=_txt(label_row.get("descricao")) or "Peca em producao",
                meta=[
                    ("Encomenda", _txt(label_row.get("encomenda")) or "-"),
                    ("OPP", _txt(label_row.get("opp")) or "-"),
                    ("Origem", _txt(label_row.get("posto_origem")) or "Corte Laser"),
                    ("Seguinte", _txt(label_row.get("proximo_posto")) or "-"),
                ],
                barcode_value=_txt(label_row.get("opp")) or "OPP-EXEMPLO",
            ),
        ),
        ("06_Etiqueta_Palete_Light_V2.pdf", lambda p: render_pallet_label(p, logo, label_row)),
        ("07_Dossier_Qualidade_Light_V2.pdf", lambda p: render_quality_dossier(p, backend, logo)),
        ("08_Historico_Materia_Prima_Light_V2.pdf", lambda p: render_history(p, logo, list(backend.material_history_rows(limit=240) or []))),
        ("09_Separacao_Material_Light_V2.pdf", lambda p: render_separation(p, logo, list(backend.material_assistant_separation_rows(horizon_days=5) or []))),
        ("10_Planeamento_Prazos_Light_V2.pdf", lambda p: render_planning(p, logo, list(backend.planning_laser_deadline_rows() or []))),
        ("11_Stock_Materia_Prima_Light_V2.pdf", lambda p: render_material_stock(p, logo, material_stock_rows)),
        ("12_Stock_Produtos_Light_V2.pdf", lambda p: render_product_stock(p, logo, product_stock_rows)),
        (
            "13_Pedido_Cotacao_Light_V2.pdf",
            lambda p: render_commercial_light(p, logo, kind="Pedido de cotacao", record=purchase_note, party=supplier, issuer=issuer),
        ),
        (
            "14_Nota_Encomenda_Light_V2.pdf",
            lambda p: render_commercial_light(p, logo, kind="Nota de encomenda", record=purchase_note, party=supplier, issuer=issuer),
        ),
        (
            "15_Orcamento_Light_V2.pdf",
            lambda p: render_commercial_light(
                p,
                logo,
                kind="Orcamento",
                record=quote,
                party=dict(quote.get("cliente") or {}),
                issuer=issuer,
            ),
        ),
    ]

    report = []
    for filename, renderer in outputs:
        target = out / filename
        renderer(target)
        pages = len(PdfReader(str(target)).pages)
        report.append(f"[OK] {filename} | {pages} pagina(s) | {target.stat().st_size} bytes")

    (out / "LEIA-ME.txt").write_text(
        "\n".join(
            [
                "PROPOSTAS PDF LUGEST LIGHT V2",
                "",
                "Estes documentos sao prototipos para avaliacao visual.",
                "Nao substituem ainda os geradores usados em producao.",
                "A folha de rota usa dados sinteticos e nao tem validade operacional.",
                "Todos os documentos usam um layout economico, sem grandes areas de cor preenchida.",
                "",
                *report,
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    print(out)
    for line in report:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
