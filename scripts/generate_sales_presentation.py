from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PIL import Image
from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


PAGE_W, PAGE_H = landscape(A4)
MARGIN = 24
NAVY = HexColor("#0B1354")
ORANGE = HexColor("#F28A1B")
RED = HexColor("#EE5A2A")
GREEN = HexColor("#1F9D73")
PALE = HexColor("#EEF4FB")
TEXT = HexColor("#18263B")
MUTED = HexColor("#667A95")
BORDER = HexColor("#C6D6EA")
SOFT_ORANGE = HexColor("#FFF3E5")
SOFT_BLUE = HexColor("#F7FAFE")


@dataclass(frozen=True)
class Callout:
    title: str
    text: str
    target_x: float
    target_y: float
    accent: Color = ORANGE


@dataclass(frozen=True)
class FeatureSlide:
    section: str
    title: str
    subtitle: str
    image_key: str
    promise: str
    callouts: tuple[Callout, ...]
    bottom_line: str


def resolve_image(image_dir: Path, stem_suffix: str) -> Path:
    matches = sorted(image_dir.glob(f"*{stem_suffix}.png"))
    if not matches:
        raise FileNotFoundError(f"Imagem nao encontrada para sufixo: {stem_suffix}")
    return matches[0]


def draw_round_box(
    c: canvas.Canvas,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: Color,
    stroke: Color | None = None,
    radius: float = 16,
) -> None:
    c.setFillColor(fill)
    c.setStrokeColor(stroke or fill)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=1 if stroke else 0)


def wrap_by_width(text: str, max_width: float, font_name: str, font_size: float) -> list[str]:
    words = text.split()
    if not words:
        return [""]
    lines: list[str] = []
    current = words[0]
    for word in words[1:]:
        trial = f"{current} {word}"
        if stringWidth(trial, font_name, font_size) <= max_width:
            current = trial
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def draw_logo(c: canvas.Canvas, logo_path: Path, x: float, y: float, w: float, h: float) -> None:
    with Image.open(logo_path) as img:
        iw, ih = img.size
    ratio = min(w / iw, h / ih)
    draw_w = iw * ratio
    draw_h = ih * ratio
    draw_x = x + (w - draw_w) / 2
    draw_y = y + (h - draw_h) / 2
    c.drawImage(ImageReader(str(logo_path)), draw_x, draw_y, width=draw_w, height=draw_h, mask="auto")


def draw_header(c: canvas.Canvas, logo_path: Path, page_no: int, total_pages: int) -> None:
    c.setFillColor(white)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(SOFT_BLUE)
    c.rect(0, PAGE_H - 48, PAGE_W, 48, fill=1, stroke=0)
    draw_logo(c, logo_path, MARGIN, PAGE_H - 44, 92, 28)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9.5)
    c.drawRightString(PAGE_W - MARGIN, PAGE_H - 28, f"Apresentacao comercial  |  {page_no}/{total_pages}")


def draw_footer(c: canvas.Canvas) -> None:
    c.setStrokeColor(BORDER)
    c.setLineWidth(0.8)
    c.line(MARGIN, 20, PAGE_W - MARGIN, 20)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9)
    c.drawString(MARGIN, 9, "LuisGEST  |  ERP industrial para producao, stock, qualidade e logistica.")


def draw_cover(c: canvas.Canvas, logo_path: Path) -> None:
    c.setFillColor(white)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(SOFT_BLUE)
    c.rect(0, PAGE_H * 0.55, PAGE_W, PAGE_H * 0.45, fill=1, stroke=0)
    draw_round_box(c, 28, 34, PAGE_W - 56, PAGE_H - 68, white, BORDER, 24)

    c.setFillColor(Color(242 / 255, 138 / 255, 27 / 255, alpha=0.18))
    c.circle(PAGE_W - 122, PAGE_H - 112, 76, fill=1, stroke=0)
    c.setFillColor(Color(11 / 255, 19 / 255, 84 / 255, alpha=0.08))
    c.circle(PAGE_W - 174, PAGE_H - 154, 110, fill=1, stroke=0)
    draw_logo(c, logo_path, 48, PAGE_H - 230, 320, 138)

    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(52, PAGE_H - 276, "Apresentacao comercial")
    c.setFont("Helvetica-Bold", 23)
    c.drawString(52, PAGE_H - 308, "LuisGEST: organizacao industrial com visibilidade real")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 12.5)
    subtitle = (
        "Uma plataforma para empresas que querem controlar melhor a proposta comercial, o stock, a producao, "
        "a qualidade e a expedicao sem depender de folhas soltas ou informacao dispersa."
    )
    y = PAGE_H - 334
    for line in wrap_by_width(subtitle, 490, "Helvetica", 12.5):
        c.drawString(52, y, line)
        y -= 17

    draw_round_box(c, 52, 88, 356, 128, SOFT_ORANGE, BORDER, 18)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(70, 186, "Porque este software merece ser testado")
    points = (
        "Liga areas que normalmente vivem separadas: comercial, producao, stock, compras e qualidade.",
        "Da visibilidade ao que esta em curso, ao que esta em risco e ao que precisa de acao.",
        "Cria base para crescer com mais rapidez, rastreabilidade e disciplina operacional.",
    )
    yy = 162
    for point in points:
        draw_bullet_line(c, point, 72, yy, 308, ORANGE)
        yy -= 28

    draw_round_box(c, 434, 88, PAGE_W - 486, 128, NAVY, None, 18)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(454, 186, "Ideal para empresas que procuram")
    lines = (
        "Menos improviso operacional",
        "Mais rapidez na resposta ao cliente",
        "Mais controlo sobre o processo completo",
    )
    yy = 156
    for line in lines:
        c.drawString(454, yy, line)
        yy -= 22


def draw_bullet_line(c: canvas.Canvas, text: str, x: float, y: float, width: float, accent: Color) -> None:
    c.setFillColor(accent)
    c.circle(x + 4, y - 4, 2.8, fill=1, stroke=0)
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10.8)
    lines = wrap_by_width(text, width - 18, "Helvetica", 10.8)
    for idx, line in enumerate(lines[:2]):
        c.drawString(x + 16, y - 8 - idx * 13, line)


def draw_screenshot(c: canvas.Canvas, image_path: Path) -> tuple[float, float, float, float]:
    box_x = 28
    box_y = 86
    box_w = 538
    box_h = 286
    c.setFillColor(Color(0, 0, 0, alpha=0.08))
    c.roundRect(box_x + 6, box_y - 6, box_w, box_h, 18, fill=1, stroke=0)
    draw_round_box(c, box_x, box_y, box_w, box_h, white, BORDER, 18)

    with Image.open(image_path) as img:
        iw, ih = img.size
    ratio = min((box_w - 16) / iw, (box_h - 16) / ih)
    draw_w = iw * ratio
    draw_h = ih * ratio
    draw_x = box_x + (box_w - draw_w) / 2
    draw_y = box_y + (box_h - draw_h) / 2
    c.drawImage(ImageReader(str(image_path)), draw_x, draw_y, width=draw_w, height=draw_h, mask="auto")
    return draw_x, draw_y, draw_w, draw_h


def draw_callout_badge(c: canvas.Canvas, number: int, x: float, y: float, accent: Color) -> None:
    c.setFillColor(white)
    c.circle(x, y, 11, fill=1, stroke=0)
    c.setFillColor(accent)
    c.circle(x, y, 8.5, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(x, y - 3.5, str(number))


def draw_arrow(c: canvas.Canvas, x1: float, y1: float, x2: float, y2: float, accent: Color) -> None:
    c.setStrokeColor(accent)
    c.setFillColor(accent)
    c.setLineWidth(1.5)
    c.line(x1, y1, x2, y2)
    import math

    angle = math.atan2(y2 - y1, x2 - x1)
    size = 7
    left = angle + 2.65
    right = angle - 2.65
    c.line(x2, y2, x2 + size * math.cos(left), y2 + size * math.sin(left))
    c.line(x2, y2, x2 + size * math.cos(right), y2 + size * math.sin(right))


def draw_callout_card(c: canvas.Canvas, idx: int, callout: Callout, x: float, y: float, w: float, h: float) -> tuple[float, float]:
    draw_round_box(c, x, y, w, h, white, BORDER, 16)
    c.setFillColor(callout.accent)
    c.rect(x, y + h - 6, w, 6, fill=1, stroke=0)
    draw_callout_badge(c, idx, x + 18, y + h - 18, callout.accent)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11.5)
    for line_idx, line in enumerate(wrap_by_width(callout.title, w - 46, "Helvetica-Bold", 11.5)[:2]):
        c.drawString(x + 34, y + h - 23 - line_idx * 13, line)
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10)
    text_y = y + h - 48
    for line_idx, line in enumerate(wrap_by_width(callout.text, w - 28, "Helvetica", 10)[:4]):
        c.drawString(x + 14, text_y - line_idx * 12.5, line)
    return x, y + h / 2


def draw_feature_slide(c: canvas.Canvas, slide: FeatureSlide, image_dir: Path) -> None:
    c.setFillColor(MUTED)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(MARGIN, PAGE_H - 76, slide.section.upper())

    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 23)
    title_y = PAGE_H - 100
    for idx, line in enumerate(wrap_by_width(slide.title, PAGE_W - 2 * MARGIN, "Helvetica-Bold", 23)[:2]):
        c.drawString(MARGIN, title_y - idx * 25, line)

    c.setFillColor(MUTED)
    c.setFont("Helvetica", 11.5)
    sub_y = PAGE_H - 128
    for idx, line in enumerate(wrap_by_width(slide.subtitle, PAGE_W - 2 * MARGIN, "Helvetica", 11.5)[:2]):
        c.drawString(MARGIN, sub_y - idx * 15, line)

    draw_round_box(c, MARGIN, 386, 266, 44, SOFT_ORANGE, BORDER, 14)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11.2)
    c.drawString(MARGIN + 14, 408, "Promessa de valor")
    c.setFillColor(TEXT)
    c.setFont("Helvetica", 10.2)
    for idx, line in enumerate(wrap_by_width(slide.promise, 230, "Helvetica", 10.2)[:2]):
        c.drawString(MARGIN + 110, 408 - idx * 12, line)

    img_x, img_y, img_w, img_h = draw_screenshot(c, resolve_image(image_dir, slide.image_key))

    sidebar_x = 584
    sidebar_w = PAGE_W - sidebar_x - 26
    card_h = 84
    card_positions = [288, 188, 88]
    anchors: list[tuple[float, float, float, float]] = []
    for idx, callout in enumerate(slide.callouts[:3], start=1):
        card_x, card_y_center = draw_callout_card(c, idx, callout, sidebar_x, card_positions[idx - 1], sidebar_w, card_h)
        target_x = img_x + callout.target_x * img_w
        target_y = img_y + callout.target_y * img_h
        draw_callout_badge(c, idx, target_x, target_y, callout.accent)
        draw_arrow(c, card_x, card_y_center, target_x - 14, target_y, callout.accent)
        anchors.append((target_x, target_y, card_x, card_y_center))

    draw_round_box(c, 28, 38, PAGE_W - 56, 34, NAVY, None, 12)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 11.5)
    c.drawString(42, 50, slide.bottom_line)


def draw_closing(c: canvas.Canvas, logo_path: Path) -> None:
    c.setFillColor(white)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    draw_round_box(c, 28, 34, PAGE_W - 56, PAGE_H - 68, SOFT_BLUE, BORDER, 24)
    draw_logo(c, logo_path, 44, PAGE_H - 128, 180, 68)

    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 27)
    c.drawString(46, PAGE_H - 176, "Porque vale a pena experimentar o LuisGEST")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 12)
    intro = (
        "Se a sua empresa sente perda de tempo entre departamentos, dificuldade em acompanhar o estado real das encomendas "
        "ou dependencia de folhas externas, este software merece uma demonstracao pratica."
    )
    y = PAGE_H - 202
    for line in wrap_by_width(intro, 690, "Helvetica", 12):
        c.drawString(46, y, line)
        y -= 16

    cards = [
        ("Mais controlo", "Sabe o que esta em curso, o que esta em risco e o que precisa de acao antes que o problema escale."),
        ("Mais rapidez", "Responde melhor ao cliente, prepara melhor a producao e reduz o tempo perdido a procurar informacao."),
        ("Mais confianca", "Trabalha com mais rastreabilidade em compras, qualidade, producao e expedicao."),
    ]
    x = 46
    card_w = (PAGE_W - 116) / 3
    for title, desc in cards:
        draw_round_box(c, x, 164, card_w, 170, white, BORDER, 18)
        c.setFillColor(ORANGE)
        c.rect(x, 308, card_w, 26, fill=1, stroke=0)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 13.5)
        c.drawString(x + 14, 316, title)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 10.6)
        yy = 286
        for line in wrap_by_width(desc, card_w - 28, "Helvetica", 10.6)[:5]:
            c.drawString(x + 14, yy, line)
            yy -= 13
        x += card_w + 12

    draw_round_box(c, 46, 68, PAGE_W - 92, 68, NAVY, None, 16)
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(64, 106, "Proximo passo recomendado")
    c.setFont("Helvetica", 11.4)
    c.drawString(64, 84, "Agende uma demonstracao orientada ao seu fluxo real: proposta, stock, planeamento, producao, qualidade e entrega.")


def build_presentation(image_dir: Path, output_pdf: Path, logo_path: Path) -> None:
    slides = [
        FeatureSlide(
            section="Visao global",
            title="Painel unico para decidir melhor e agir mais cedo",
            subtitle="O dashboard concentra indicadores e atalhos de trabalho para que a direcao e a operacao leiam rapidamente o estado real da empresa.",
            image_key="170256",
            promise="Uma unica vista para perceber o que esta a acontecer e o que exige resposta imediata.",
            callouts=(
                Callout("Leitura imediata", "Os indicadores ajudam a perceber rapidamente o que esta em curso, em risco ou a aguardar acao.", 0.26, 0.28),
                Callout("Atalhos operacionais", "Acoes chave como abrir encomenda, transporte ou exportacao ficam acessiveis sem navegar por varios menus.", 0.82, 0.18, RED),
                Callout("Analise por area", "O painel separa operacao e stock/compras para apoiar reunioes diarias e prioridades da gestao.", 0.17, 0.62, GREEN),
            ),
            bottom_line="Resultado esperado: menos tempo a procurar informacao e mais rapidez a decidir onde atuar.",
        ),
        FeatureSlide(
            section="Orcamentacao",
            title="Transforme propostas em encomendas com muito menos friccao",
            subtitle="O modulo de orcamentos liga cliente, configuracao tecnica, linhas, operacoes e total financeiro num unico fluxo comercial.",
            image_key="165943",
            promise="Responder mais rapido ao cliente sem perder coerencia entre o que se promete e o que se vai fabricar.",
            callouts=(
                Callout("Fluxo comercial completo", "Aprovacao, envio, PDF e conversao em encomenda estao no mesmo registo.", 0.36, 0.18),
                Callout("Resumo financeiro visivel", "O comercial ganha controlo imediato sobre linhas, transporte, desconto e total.", 0.92, 0.42, RED),
                Callout("Referencias tecnicas ligadas", "A proposta pode incluir desenho, nesting, configuracoes e preparacao de producao.", 0.63, 0.78, GREEN),
            ),
            bottom_line="Resultado esperado: mais velocidade comercial e menos erros na passagem da proposta para a producao.",
        ),
        FeatureSlide(
            section="Planeamento e producao",
            title="Planeie a semana com visibilidade sobre carga e prioridades",
            subtitle="O planeamento semanal ajuda a distribuir trabalho por operacao e recurso, reduzindo improviso e melhorando o cumprimento de prazos.",
            image_key="170020",
            promise="Dar ao responsavel de producao uma vista clara para organizar a semana antes do problema chegar ao chao de fabrica.",
            callouts=(
                Callout("Controlo do horizonte", "Ferramentas de navegacao por semana e dia ajudam a ajustar carga e janelas de trabalho.", 0.59, 0.18),
                Callout("Leitura de carga", "Os blocos de resumo mostram rapidamente encomendas pendentes, carga semanal e blocos fechados.", 0.58, 0.33, GREEN),
                Callout("Quadro visual por horario", "A grelha semanal facilita o encaixe das encomendas e a identificacao de folgas ou conflitos.", 0.67, 0.69, RED),
            ),
            bottom_line="Resultado esperado: mais previsibilidade, melhor sequenciacao e menos replaneamento em cima da hora.",
        ),
        FeatureSlide(
            section="Compras e logistica",
            title="Compras, fornecedores e notas de encomenda no mesmo ecossistema",
            subtitle="O software organiza pedidos a fornecedores, detalhe documental e controlo das linhas para apoiar um processo de compra mais disciplinado.",
            image_key="170146",
            promise="Ganhar organizacao na relacao com fornecedores e reduzir dependencia de email, folhas e controlo manual.",
            callouts=(
                Callout("Acoes centrais de compra", "Criar, aprovar, gerar documentos e enviar encomendas fica concentrado no topo do registo.", 0.33, 0.20),
                Callout("Resumo do documento", "Entrega, contacto, observacoes e total estao visiveis para reduzir falhas de comunicacao.", 0.82, 0.40, RED),
                Callout("Linhas bem estruturadas", "Cada artigo comprado fica controlado por codigo, material, quantidade, fornecedor e entrega.", 0.42, 0.74, GREEN),
            ),
            bottom_line="Resultado esperado: compras mais organizadas, menos omissoes e melhor ligacao entre necessidade e fornecedor.",
        ),
        FeatureSlide(
            section="Qualidade",
            title="Rastreabilidade e controlo de qualidade sem sair do sistema",
            subtitle="Rececao, inspecao, nao conformidades e evidencias passam a viver no mesmo ambiente do resto da operacao.",
            image_key="170219",
            promise="Ter mais confianca no processo e criar uma base de trabalho mais solida para controlo interno e auditoria.",
            callouts=(
                Callout("Indicadores de qualidade", "Abertas, fora de prazo, ligadas a fornecedor ou material bloqueado ficam logo visiveis.", 0.47, 0.24),
                Callout("Acoes imediatas", "Avaliar, aprovar ou rejeitar deixa de depender de registos paralelos e mensagens soltas.", 0.88, 0.40, RED),
                Callout("Lista operacional", "O detalhe por referencia, lote, fornecedor e estado ajuda a tratar rececoes com metodo.", 0.55, 0.62, GREEN),
            ),
            bottom_line="Resultado esperado: mais rigor, mais rastreabilidade e menos risco de decisoes sem evidencia.",
        ),
    ]

    total_pages = len(slides) + 2
    c = canvas.Canvas(str(output_pdf), pagesize=landscape(A4))
    c.setTitle("LuisGEST - Apresentacao Comercial")

    draw_cover(c, logo_path)
    c.showPage()

    for page_no, slide in enumerate(slides, start=2):
        draw_header(c, logo_path, page_no, total_pages)
        draw_feature_slide(c, slide, image_dir)
        draw_footer(c)
        c.showPage()

    draw_closing(c, logo_path)
    draw_footer(c)
    c.save()


def main() -> None:
    image_dir = Path(r"C:\Users\engenharia\Desktop\Prin ts")
    output_pdf = image_dir / "LuisGEST - Apresentacao Comercial.pdf"
    logo_path = Path(r"C:\Users\engenharia\VSCodeProjects\teste\Logos\logo.png")
    build_presentation(image_dir, output_pdf, logo_path)
    print(output_pdf)


if __name__ == "__main__":
    main()
