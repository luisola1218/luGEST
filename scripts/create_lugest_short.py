from __future__ import annotations

import argparse
import asyncio
import math
import os
import subprocess
import wave
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageOps

from create_lugest_promo import font, rounded_mask


WIDTH = 1080
HEIGHT = 1920
FPS = 30
OFF_WHITE = "#F5F6F3"
DARKER = "#17222D"
DARK = "#2F3B46"
MUTED = "#647181"
ORANGE = "#F4690A"
GREEN = "#69C80E"
LINE = "#D8DDD9"


def run(command: list[str]) -> None:
    subprocess.run(command, check=True)


def probe_duration(ffmpeg: Path, media: Path) -> float:
    result = subprocess.run(
        [str(ffmpeg), "-i", str(media)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )
    import re

    match = re.search(r"Duration:\s*(\d+):(\d+):([\d.]+)", result.stderr)
    if not match:
        raise RuntimeError(f"Não foi possível determinar a duração de {media}")
    return int(match.group(1)) * 3600 + int(match.group(2)) * 60 + float(match.group(3))


def fit_image(path: Path, size: tuple[int, int]) -> Image.Image:
    image = Image.open(path).convert("RGB")
    image.thumbnail(size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", size, "#FFFFFF")
    canvas.paste(image, ((size[0] - image.width) // 2, (size[1] - image.height) // 2))
    return canvas


def cover_image(path: Path, size: tuple[int, int]) -> Image.Image:
    image = Image.open(path).convert("RGB")
    scale = max(size[0] / image.width, size[1] / image.height)
    image = image.resize(
        (round(image.width * scale), round(image.height * scale)),
        Image.Resampling.LANCZOS,
    )
    left = (image.width - size[0]) // 2
    top = (image.height - size[1]) // 2
    return image.crop((left, top, left + size[0], top + size[1]))


def paste_card(
    canvas: Image.Image,
    screenshot: Path,
    box: tuple[int, int, int, int],
    radius: int = 30,
) -> None:
    x1, y1, x2, y2 = box
    width, height = x2 - x1, y2 - y1
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(
        (x1 + 10, y1 + 18, x2 + 10, y2 + 18),
        radius,
        fill=(23, 34, 45, 42),
    )
    canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(14)))
    content = fit_image(screenshot, (width, height)).convert("RGBA")
    content.putalpha(rounded_mask((width, height), radius))
    canvas.alpha_composite(content, (x1, y1))
    ImageDraw.Draw(canvas).rounded_rectangle(box, radius, outline="#C8CFCA", width=2)


def logo_mark(canvas: Image.Image, logo: Image.Image, x: int = 56, y: int = 50) -> None:
    mark = logo.copy()
    mark.thumbnail((150, 150), Image.Resampling.LANCZOS)
    canvas.alpha_composite(mark, (x, y))


def transparent_wordmark(path: Path) -> Image.Image:
    """Recorta o logótipo completo e remove apenas o fundo quase branco."""
    source = Image.open(path).convert("RGBA")
    rgb = source.convert("RGB")
    bbox = Image.eval(
        rgb.convert("L"),
        lambda value: 255 if value < 246 else 0,
    ).getbbox()
    if bbox:
        source = source.crop(bbox)
    pixels = source.load()
    for y in range(source.height):
        for x in range(source.width):
            red, green, blue, _ = pixels[x, y]
            whiteness = min(red, green, blue)
            alpha = max(0, min(255, (245 - whiteness) * 18))
            pixels[x, y] = (red, green, blue, alpha)
    return source


def transparent_partner_logo(path: Path) -> Image.Image:
    """Recorta o logótipo sem halos, preservando o símbolo preto dentro do círculo."""
    source = Image.open(path).convert("RGB")
    rgb = np.asarray(source, dtype=np.float32)
    height, width = rgb.shape[:2]

    # O PNG original vem aplicado sobre preto. Fora do emblema circular, a
    # luminância separa com precisão o lettering do fundo e produz um alfa
    # suave, sem os resíduos que o flood-fill deixava dentro das letras.
    peak = rgb.max(axis=2)
    alpha = np.clip((peak - 4.0) / 20.0, 0.0, 1.0)

    # O monograma preto está dentro do disco vermelho e deve permanecer
    # totalmente opaco. Localizamos apenas o grande círculo da esquerda,
    # ignorando o lettering vermelho de "STUDIO SL".
    red = rgb[:, :, 0]
    green = rgb[:, :, 1]
    blue = rgb[:, :, 2]
    x_grid = np.arange(width, dtype=np.float32)[None, :]
    red_circle = (red > 100) & (red > green * 1.8) & (red > blue * 1.5) & (x_grid < width * 0.36)
    ys, xs = np.where(red_circle)
    if xs.size and ys.size:
        left, right = float(xs.min()), float(xs.max())
        top, bottom = float(ys.min()), float(ys.max())
        center_x = (left + right) / 2.0
        center_y = (top + bottom) / 2.0
        radius_x = max(1.0, (right - left) / 2.0)
        radius_y = max(1.0, (bottom - top) / 2.0)
        yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
        distance = np.sqrt(((xx - center_x) / radius_x) ** 2 + ((yy - center_y) / radius_y) ** 2)
        disk_alpha = np.clip((1.012 - distance) / 0.024, 0.0, 1.0)
        circle_zone = xx <= (right + 18.0)
        alpha = np.where(circle_zone, disk_alpha, alpha)

    # A assinatura vermelha do ficheiro original possui um sombreado preto
    # pensado para fundo escuro. Num cartão branco esse sombreado parece ruído;
    # isolamos apenas o traço vermelho da marca e conservamos o antialias.
    red_lettering = (red > 65) & (red > green * 1.55) & (red > blue * 1.35) & (x_grid > width * 0.34)
    red_ys, _ = np.where(red_lettering)
    if red_ys.size:
        yy, xx = np.mgrid[0:height, 0:width].astype(np.float32)
        lettering_zone = (xx > width * 0.34) & (yy >= float(red_ys.min() - 12))
        red_strength = red - np.maximum(green, blue)
        red_alpha = np.clip((red_strength - 5.0) / 38.0, 0.0, 1.0)
        alpha = np.where(lettering_zone, red_alpha, alpha)

    # Compensa a mistura dos píxeis de contorno com o antigo fundo preto.
    # Assim, o vermelho e o lettering mantêm-se sólidos quando aplicados no
    # cartão branco, mesmo depois de reduzidos para o vídeo vertical.
    soft_edge = (alpha > 0.04) & (alpha < 0.98)
    rgb[soft_edge] = np.clip(rgb[soft_edge] / alpha[soft_edge, None], 0.0, 255.0)

    alpha_u8 = np.rint(alpha * 255.0).astype(np.uint8)
    rgba = np.dstack((rgb.astype(np.uint8), alpha_u8))
    cleaned = Image.fromarray(rgba, "RGBA")
    bbox = cleaned.getchannel("A").getbbox()
    if bbox:
        cleaned = cleaned.crop((max(0, bbox[0] - 4), max(0, bbox[1] - 4), min(width, bbox[2] + 4), min(height, bbox[3] + 4)))

    # A redução final beneficia de uma passagem de nitidez muito subtil.
    return cleaned.filter(ImageFilter.UnsharpMask(radius=1.0, percent=115, threshold=3))


def add_partner_footer(
    slide: Image.Image,
    partner: Image.Image,
    language: str,
) -> Image.Image:
    """Assinatura discreta nas páginas de produto, sem competir com o luGEST."""
    canvas = slide.copy()
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 1792, WIDTH, HEIGHT), fill=OFF_WHITE)
    draw.line((56, 1792, 1024, 1792), fill=LINE, width=2)
    draw.text(
        (56, 1850),
        "luGEST  |  SOFTWARE ERP INDUSTRIAL",
        font=font(22, True),
        fill=MUTED,
    )
    label = "REPRESENTANTE OFICIAL"
    draw.text((1022, 1814), label, font=font(15, True), fill="#7B3030", anchor="ra")
    mark = partner.copy()
    mark.thumbnail((300, 64), Image.Resampling.LANCZOS)
    canvas.alpha_composite(mark, (1022 - mark.width, 1841))
    return canvas


def paste_wordmark(
    canvas: Image.Image,
    wordmark: Image.Image,
    max_size: tuple[int, int],
    center: tuple[int, int],
) -> None:
    mark = wordmark.copy()
    mark.thumbnail(max_size, Image.Resampling.LANCZOS)
    canvas.alpha_composite(
        mark,
        (center[0] - mark.width // 2, center[1] - mark.height // 2),
    )


def wrap_text(draw: ImageDraw.ImageDraw, text: str, text_font, max_width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if draw.textbbox((0, 0), candidate, font=text_font)[2] <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def content_slide(
    logo: Image.Image,
    screenshot: Path,
    kicker: str,
    title: str,
    subtitle: str,
    benefit: str,
) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 72), kicker.upper(), font=font(27, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=LINE, width=2)

    title_font = font(61, True)
    title_lines = wrap_text(draw, title, title_font, 960)
    title_y = 245
    for line in title_lines:
        draw.text((56, title_y), line, font=title_font, fill=DARKER)
        title_y += 72

    subtitle_y = title_y + 22
    for line in wrap_text(draw, subtitle, font(34), 950):
        draw.text((58, subtitle_y), line, font=font(34), fill=MUTED)
        subtitle_y += 45

    card_top = max(570, subtitle_y + 55)
    paste_card(canvas, screenshot, (56, card_top, 1024, 1430))

    draw.rounded_rectangle((56, 1495, 1024, 1745), 28, fill=DARKER)
    draw.rectangle((56, 1495, 70, 1745), fill=ORANGE)
    benefit_font = font(38, True)
    lines = wrap_text(draw, benefit, benefit_font, 880)
    line_height = 49
    start_y = 1620 - (len(lines) * line_height) // 2
    for line in lines:
        draw.text((105, start_y), line, font=benefit_font, fill="#FFFFFF")
        start_y += line_height

    draw.text((56, 1838), "luGEST  |  SOFTWARE ERP INDUSTRIAL", font=font(23, True), fill=MUTED)
    return canvas


def stock_slide(logo: Image.Image, materials: Path, purchasing: Path) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 72), "COMPRAS, MATÉRIA-PRIMA E STOCKS", font=font(25, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=LINE, width=2)
    draw.text((56, 245), "Nada falha. O stock acompanha.", font=font(53, True), fill=DARKER)
    draw.text((58, 330), "Necessidades, compras, entradas, reservas e valorização ligadas.", font=font(29), fill=MUTED)
    paste_card(canvas, purchasing, (56, 430, 1024, 930))
    paste_card(canvas, materials, (56, 970, 1024, 1470))
    draw.rounded_rectangle((56, 1520, 1024, 1780), 28, fill=GREEN)
    draw.text((540, 1605), "COMPRAR CERTO. ATUALIZAR AUTOMATICAMENTE.", font=font(31, True), fill=DARKER, anchor="mm")
    draw.text((540, 1670), "Da necessidade à receção, com stock sempre atual.", font=font(29), fill=DARKER, anchor="mm")
    draw.text((540, 1720), "Mais controlo. Menos ruturas.", font=font(29, True), fill=DARKER, anchor="mm")
    draw.text((56, 1840), "luGEST  |  SOFTWARE ERP INDUSTRIAL", font=font(23, True), fill=MUTED)
    return canvas


def cinematic_factory_slide(
    logo: Image.Image,
    background: Path,
    kicker: str,
    title: str,
    subtitle: str,
    button: str | None = None,
    clicked: bool = False,
    hero_symbol: bool = False,
) -> Image.Image:
    canvas = cover_image(background, (WIDTH, HEIGHT)).convert("RGBA")
    shade = Image.new("RGBA", canvas.size, (10, 18, 27, 132))
    gradient = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    gradient_draw = ImageDraw.Draw(gradient)
    for y in range(HEIGHT):
        alpha = min(205, max(25, int(215 - y * 0.07)))
        gradient_draw.line((0, y, WIDTH, y), fill=(10, 18, 27, alpha))
    canvas.alpha_composite(shade)
    canvas.alpha_composite(gradient)
    draw = ImageDraw.Draw(canvas)

    logo_mark(canvas, logo, 58, 54)
    draw.text((230, 78), kicker.upper(), font=font(26, True), fill=ORANGE)
    draw.line((58, 195, 1022, 195), fill=(255, 255, 255, 95), width=2)

    title_font = font(72, True)
    title_y = 285
    for line in wrap_text(draw, title, title_font, 930):
        draw.text((58, title_y), line, font=title_font, fill="#FFFFFF")
        title_y += 86
    subtitle_y = title_y + 34
    for line in wrap_text(draw, subtitle, font(36), 890):
        draw.text((60, subtitle_y), line, font=font(36), fill="#DDE4E8")
        subtitle_y += 50

    if hero_symbol:
        # Assinatura visual de entrada: o símbolo ocupa deliberadamente a
        # zona negativa da composição, sem tapar o gestor nem a mensagem.
        hero_box = (310, 760, 770, 1220)
        glow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        ImageDraw.Draw(glow).rounded_rectangle(
            hero_box,
            56,
            fill=(244, 105, 10, 62),
        )
        canvas.alpha_composite(glow.filter(ImageFilter.GaussianBlur(34)))
        draw.rounded_rectangle(
            hero_box,
            56,
            fill=(10, 18, 27, 205),
            outline=(255, 255, 255, 55),
            width=2,
        )
        hero = logo.copy()
        hero.thumbnail((350, 350), Image.Resampling.LANCZOS)
        canvas.alpha_composite(
            hero,
            (
                (hero_box[0] + hero_box[2] - hero.width) // 2,
                (hero_box[1] + hero_box[3] - hero.height) // 2,
            ),
        )

    if button:
        button_box = (120, 1510, 960, 1645)
        shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
        ImageDraw.Draw(shadow).rounded_rectangle(
            (130, 1524, 970, 1659), 67, fill=(0, 0, 0, 90)
        )
        canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(14)))
        draw.rounded_rectangle(button_box, 67, fill=GREEN)
        draw.text(
            (540, 1577),
            button.upper(),
            font=font(35, True),
            fill=DARKER,
            anchor="mm",
        )
        cursor_x, cursor_y = 850, 1605
        draw.polygon(
            (
                (cursor_x, cursor_y),
                (cursor_x + 16, cursor_y + 58),
                (cursor_x + 31, cursor_y + 40),
                (cursor_x + 57, cursor_y + 66),
                (cursor_x + 71, cursor_y + 52),
                (cursor_x + 45, cursor_y + 27),
                (cursor_x + 65, cursor_y + 15),
            ),
            fill="#FFFFFF",
            outline=DARKER,
        )
        if clicked:
            for radius, alpha in ((46, 220), (72, 135), (98, 70)):
                draw.ellipse(
                    (
                        cursor_x - radius,
                        cursor_y - radius,
                        cursor_x + radius,
                        cursor_y + radius,
                    ),
                    outline=(255, 255, 255, alpha),
                    width=5,
                )

    draw.text(
        (58, 1845),
        "luGEST  |  SOFTWARE ERP INDUSTRIAL",
        font=font(23, True),
        fill="#D7DEE3",
    )
    return canvas


def operations_slide(
    logo: Image.Image,
    background: Path,
    planning: Path,
) -> Image.Image:
    canvas = cover_image(background, (WIDTH, HEIGHT)).convert("RGBA")
    canvas.alpha_composite(Image.new("RGBA", canvas.size, (10, 18, 27, 108)))
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 75), "CONTROLO OPERACIONAL", font=font(27, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=(255, 255, 255, 105), width=2)
    draw.text((56, 245), "Decidir antes do problema.", font=font(61, True), fill="#FFFFFF")
    draw.multiline_text(
        (58, 340),
        "Planeamento, compras e stock\nligados à realidade da fábrica.",
        font=font(34),
        fill="#DDE4E8",
        spacing=12,
    )
    paste_card(canvas, planning, (56, 600, 1024, 1260))
    chips = (
        ("PLANEAMENTO", 56, 1330, 351),
        ("COMPRAS", 375, 1330, 670),
        ("STOCK", 694, 1330, 1024),
    )
    for text, x1, y1, x2 in chips:
        draw.rounded_rectangle((x1, y1, x2, 1435), 26, fill=(23, 34, 45, 225))
        draw.text(((x1 + x2) // 2, 1382), text, font=font(24, True), fill="#FFFFFF", anchor="mm")
    draw.rounded_rectangle((56, 1510, 1024, 1765), 28, fill=GREEN)
    draw.text(
        (540, 1595),
        "VISIBILIDADE PARA AGIR.",
        font=font(38, True),
        fill=DARKER,
        anchor="mm",
    )
    draw.text(
        (540, 1670),
        "Controlo para orientar resultados.",
        font=font(31),
        fill=DARKER,
        anchor="mm",
    )
    draw.text((56, 1840), "luGEST  |  SOFTWARE ERP INDUSTRIAL", font=font(23, True), fill="#D7DEE3")
    return canvas


def brand_intro_slide(
    wordmark: Image.Image,
    background: Path,
) -> Image.Image:
    canvas = cover_image(background, (WIDTH, HEIGHT)).convert("RGBA")
    canvas.alpha_composite(Image.new("RGBA", canvas.size, (8, 15, 23, 106)))
    gradient = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    gradient_draw = ImageDraw.Draw(gradient)
    for y in range(1100):
        alpha = max(10, 188 - int(y * 0.13))
        gradient_draw.line((0, y, WIDTH, y), fill=(8, 15, 23, alpha))
    canvas.alpha_composite(gradient)
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle(
        (50, 92, 1030, 735),
        42,
        fill=(11, 20, 29, 198),
        outline=(255, 255, 255, 42),
        width=2,
    )
    paste_wordmark(canvas, wordmark, (900, 410), (540, 330))
    draw.line((140, 545, 940, 545), fill=(244, 105, 10, 205), width=4)
    draw.text((540, 620), "CONTROLO PARA DECIDIR", font=font(38, True), fill="#FFFFFF", anchor="mm")
    draw.text((540, 680), "TRANSPARÊNCIA PARA CRESCER", font=font(29, True), fill="#F8A15D", anchor="mm")
    draw.text((58, 1835), "SOFTWARE ERP INDUSTRIAL", font=font(24, True), fill="#E2E7EA")
    return canvas


def finance_slide(
    logo: Image.Image,
    dashboard: Path,
) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 72), "CONTROLO FINANCEIRO", font=font(27, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=LINE, width=2)
    draw.text((56, 245), "Sabe quanto vale a sua empresa?", font=font(54, True), fill=DARKER)
    draw.multiline_text(
        (58, 340),
        "Stock valorizado, compras, vendas,\ncompromissos e margens numa única visão.",
        font=font(31),
        fill=MUTED,
        spacing=10,
    )
    paste_card(canvas, dashboard, (56, 535, 1024, 1315))
    metrics = (
        ("STOCK", "VALORIZADO"),
        ("COMPRAS", "CONTROLADAS"),
        ("MARGENS", "VISÍVEIS"),
    )
    for (label, value), x in zip(metrics, (56, 381, 706)):
        draw.rounded_rectangle((x, 1380, x + 294, 1545), 24, fill="#FFFFFF", outline=LINE, width=2)
        draw.text((x + 147, 1427), label, font=font(22, True), fill=ORANGE, anchor="mm")
        draw.text((x + 147, 1490), value, font=font(25, True), fill=DARKER, anchor="mm")
    draw.rounded_rectangle((56, 1600, 1024, 1775), 28, fill=DARKER)
    draw.text(
        (540, 1663),
        "CONTROLO FINANCEIRO PARA DECIDIR COM CONFIANÇA.",
        font=font(29, True),
        fill="#FFFFFF",
        anchor="mm",
    )
    draw.text(
        (540, 1718),
        "Uma leitura clara do presente. Uma base sólida para crescer.",
        font=font(26),
        fill="#D8E0E5",
        anchor="mm",
    )
    draw.text((56, 1840), "luGEST  |  SOFTWARE ERP INDUSTRIAL", font=font(23, True), fill=MUTED)
    return canvas


def adaptable_slide(
    logo: Image.Image,
    background: Path,
) -> Image.Image:
    canvas = cover_image(background, (WIDTH, HEIGHT)).convert("RGBA")
    canvas.alpha_composite(Image.new("RGBA", canvas.size, (7, 14, 22, 122)))
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 75), "FEITO À MEDIDA DA SUA EMPRESA", font=font(25, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=(255, 255, 255, 105), width=2)
    draw.rounded_rectangle((46, 235, 1018, 670), 34, fill=(10, 19, 28, 214))
    draw.text((76, 285), "O software adapta-se.", font=font(62, True), fill="#FFFFFF")
    draw.text((76, 370), "A sua empresa não tem de se adaptar.", font=font(37, True), fill="#F8A15D")
    body = (
        "Fluxos, documentos, operações, permissões e relatórios "
        "configurados à realidade de cada cliente."
    )
    y = 465
    for line in wrap_text(draw, body, font(31), 870):
        draw.text((78, y), line, font=font(31), fill="#DDE4E8")
        y += 45
    cards = (
        ("MODULAR", "Ative apenas o que precisa."),
        ("CONFIGURÁVEL", "Acompanhe os seus processos."),
        ("EVOLUTIVO", "Cresce com a empresa."),
    )
    y = 1120
    for title, text in cards:
        draw.rounded_rectangle((70, y, 1010, y + 155), 28, fill=(245, 246, 243, 232))
        draw.rectangle((70, y, 83, y + 155), fill=GREEN)
        draw.text((115, y + 32), title, font=font(27, True), fill=DARKER)
        draw.text((115, y + 84), text, font=font(27), fill=MUTED)
        y += 182
    draw.text((58, 1838), "luGEST  |  SOFTWARE ERP INDUSTRIAL", font=font(23, True), fill="#E2E7EA")
    return canvas


def cover_slide(logo: Image.Image) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    for index in range(14):
        inset = index * 28
        draw.ellipse((440 - inset, -140 - inset, 1450 + inset, 870 + inset), outline=ORANGE, width=3)
    mark = logo.copy()
    mark.thumbnail((430, 430), Image.Resampling.LANCZOS)
    canvas.alpha_composite(mark, ((WIDTH - mark.width) // 2, 240))
    draw.text((540, 765), "luGEST", font=font(116, True), fill="#FFFFFF", anchor="mm")
    draw.text((540, 880), "SOFTWARE ERP INDUSTRIAL", font=font(29, True), fill=ORANGE, anchor="mm")
    draw.multiline_text(
        (540, 1120),
        "Controlo para decidir.\nTransparência para crescer.",
        font=font(57, True),
        fill="#FFFFFF",
        anchor="mm",
        align="center",
        spacing=18,
    )
    draw.rounded_rectangle((155, 1455, 925, 1555), 50, fill=GREEN)
    draw.text((540, 1505), "SOFTWARE ERP INDUSTRIAL", font=font(31, True), fill=DARKER, anchor="mm")
    return canvas


def clean_brand_intro_slide(wordmark: Image.Image) -> Image.Image:
    """Abertura de marca limpa, sem fotografia ou elementos gerados."""
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 0, WIDTH, 18), fill=ORANGE)
    draw.rectangle((0, HEIGHT - 18, WIDTH, HEIGHT), fill=GREEN)
    paste_wordmark(canvas, wordmark, (900, 420), (540, 600))
    draw.line((150, 900, 930, 900), fill=ORANGE, width=5)
    draw.text(
        (540, 1055),
        "CONTROLO PARA DECIDIR",
        font=font(55, True),
        fill="#FFFFFF",
        anchor="mm",
    )
    draw.text(
        (540, 1145),
        "TRANSPARÊNCIA PARA CRESCER",
        font=font(37, True),
        fill="#F58A42",
        anchor="mm",
    )
    draw.rounded_rectangle((175, 1425, 905, 1545), 24, outline="#FFFFFF", width=2)
    draw.text(
        (540, 1485),
        "SOFTWARE ERP INDUSTRIAL",
        font=font(31, True),
        fill="#FFFFFF",
        anchor="mm",
    )
    return canvas


def _ease_out_cubic(value: float) -> float:
    value = max(0.0, min(1.0, value))
    return 1.0 - (1.0 - value) ** 3


def _ease_in_out(value: float) -> float:
    value = max(0.0, min(1.0, value))
    return value * value * (3.0 - 2.0 * value)


def _silver_brand_background() -> Image.Image:
    """Fundo metálico limpo, recriado de raiz e estável entre fotogramas."""
    yy, xx = np.mgrid[0:HEIGHT, 0:WIDTH]
    distance = np.sqrt(((xx - WIDTH * 0.50) / (WIDTH * 0.78)) ** 2 + ((yy - HEIGHT * 0.38) / (HEIGHT * 0.72)) ** 2)
    light = np.clip(246.0 - distance * 52.0, 184.0, 246.0)
    blue = np.clip(light + 7.0, 190.0, 251.0)
    rgb = np.dstack((light, light + 2.0, blue)).astype(np.uint8)
    canvas = Image.fromarray(rgb, "RGB").convert("RGBA")
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 0, WIDTH, 14), fill=ORANGE)
    draw.rectangle((0, HEIGHT - 14, WIDTH, HEIGHT), fill=DARKER)
    for y in range(230, HEIGHT - 180, 110):
        draw.line((90, y, WIDTH - 90, y), fill=(255, 255, 255, 26), width=1)
    return canvas


def render_brand_intro(
    ffmpeg: Path,
    output: Path,
    work: Path,
    wordmark: Image.Image,
    symbol: Image.Image,
    duration: float,
    language: str,
) -> None:
    """Animação original inspirada na linguagem industrial do vídeo de referência."""
    frames_dir = work / "brand_intro_frames"
    frames_dir.mkdir(parents=True, exist_ok=True)
    base = _silver_brand_background()
    full_mark = wordmark.copy()
    full_mark.thumbnail((930, 500), Image.Resampling.LANCZOS)
    symbol_mark = symbol.copy()
    symbol_mark.thumbnail((300, 300), Image.Resampling.LANCZOS)
    mark_x = (WIDTH - full_mark.width) // 2
    mark_y = 545

    reflection = ImageOps.flip(full_mark)
    reflection_alpha = reflection.getchannel("A")
    gradient = Image.new("L", reflection.size, 0)
    gradient_pixels = gradient.load()
    for y in range(reflection.height):
        fade = int(58 * max(0.0, 1.0 - y / max(1, reflection.height - 1)) ** 2)
        for x in range(reflection.width):
            gradient_pixels[x, y] = min(fade, reflection_alpha.getpixel((x, y)))
    reflection.putalpha(gradient)
    reflection = reflection.filter(ImageFilter.GaussianBlur(1.2))

    if language == "es":
        intro_title = "CONTROL PARA DECIDIR"
        intro_subtitle = "TRANSPARENCIA PARA CRECER"
    else:
        intro_title = "CONTROLO PARA DECIDIR"
        intro_subtitle = "TRANSPARÊNCIA PARA CRESCER"

    total_frames = max(1, round(duration * FPS))
    for frame_index in range(total_frames):
        t = frame_index / FPS
        frame = base.copy()
        draw = ImageDraw.Draw(frame)

        # O símbolo entra primeiro com uma expansão única e contínua.
        symbol_progress = _ease_out_cubic((t - 0.18) / 0.92)
        if symbol_progress > 0:
            symbol_exit = _ease_in_out((t - 0.72) / 0.72)
            symbol_opacity = min(1.0, symbol_progress * 1.45) * (1.0 - symbol_exit)
            scale = 0.72 + 0.28 * symbol_progress
            size = (
                max(1, round(symbol_mark.width * scale)),
                max(1, round(symbol_mark.height * scale)),
            )
            animated_symbol = symbol_mark.resize(size, Image.Resampling.LANCZOS)
            animated_symbol.putalpha(
                animated_symbol.getchannel("A").point(
                    lambda alpha: round(alpha * symbol_opacity)
                )
            )
            frame.alpha_composite(
                animated_symbol,
                ((WIDTH - size[0]) // 2, 435 + round((1.0 - symbol_progress) * 44)),
            )

        # A assinatura completa é revelada da esquerda para a direita.
        reveal = _ease_in_out((t - 0.88) / 1.25)
        if reveal > 0:
            visible_width = max(1, round(full_mark.width * reveal))
            visible = full_mark.crop((0, 0, visible_width, full_mark.height))
            visible.putalpha(
                visible.getchannel("A").point(
                    lambda alpha: round(alpha * min(1.0, reveal * 1.6))
                )
            )
            frame.alpha_composite(visible, (mark_x, mark_y))
            reflected_visible = reflection.crop((0, 0, visible_width, reflection.height))
            frame.alpha_composite(reflected_visible, (mark_x, mark_y + full_mark.height + 22))

        # Varredura de luz laranja sem reamostrar o fundo, evitando vibrações.
        sweep = (t - 1.55) / 1.15
        if 0.0 <= sweep <= 1.0:
            sweep_x = round(-180 + sweep * (WIDTH + 360))
            beam = Image.new("RGBA", frame.size, (0, 0, 0, 0))
            beam_draw = ImageDraw.Draw(beam)
            beam_draw.polygon(
                (
                    (sweep_x - 95, 430),
                    (sweep_x + 35, 430),
                    (sweep_x + 190, 1040),
                    (sweep_x + 60, 1040),
                ),
                fill=(244, 105, 10, 36),
            )
            frame.alpha_composite(beam.filter(ImageFilter.GaussianBlur(26)))

        text_alpha = round(255 * _ease_in_out((t - 2.15) / 0.75))
        if text_alpha > 0:
            draw.text(
                (540, 1115),
                intro_title,
                font=font(53, True),
                fill=(23, 34, 45, text_alpha),
                anchor="mm",
            )
            draw.text(
                (540, 1205),
                intro_subtitle,
                font=font(32, True),
                fill=(244, 105, 10, text_alpha),
                anchor="mm",
            )
            line_width = round(760 * _ease_out_cubic((t - 2.35) / 0.8))
            draw.line(
                (540 - line_width // 2, 1305, 540 + line_width // 2, 1305),
                fill=(23, 34, 45, min(150, text_alpha)),
                width=3,
            )
            draw.text(
                (540, 1450),
                "SOFTWARE ERP INDUSTRIAL",
                font=font(29, True),
                fill=(47, 59, 70, text_alpha),
                anchor="mm",
            )

        frame.convert("RGB").save(frames_dir / f"frame_{frame_index:04d}.jpg", quality=94, subsampling=0)

    run(
        [
            str(ffmpeg),
            "-y",
            "-framerate",
            str(FPS),
            "-i",
            str(frames_dir / "frame_%04d.jpg"),
            "-t",
            f"{duration:.2f}",
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-crf",
            "17",
            "-pix_fmt",
            "yuv420p",
            str(output),
        ]
    )


def closing_slide(logo: Image.Image) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    for index in range(8):
        inset = index * 38
        draw.ellipse(
            (270 - inset, 105 - inset, 810 + inset, 645 + inset),
            outline=(244, 105, 10, max(35, 180 - index * 18)),
            width=4,
        )
    mark = logo.copy()
    mark.thumbnail((340, 340), Image.Resampling.LANCZOS)
    canvas.alpha_composite(mark, ((WIDTH - mark.width) // 2, 235))
    draw.multiline_text(
        (540, 850),
        "Da proposta\nà produção.",
        font=font(84, True),
        fill="#FFFFFF",
        anchor="mm",
        align="center",
        spacing=15,
    )
    draw.text((540, 1115), "Toda a operação sob controlo.", font=font(42), fill="#CBD2D7", anchor="mm")
    draw.rounded_rectangle((125, 1390, 955, 1515), 62, fill=ORANGE)
    draw.text((540, 1452), "AGENDE UMA DEMONSTRAÇÃO", font=font(32, True), fill="#FFFFFF", anchor="mm")
    draw.text((540, 1690), "luGEST", font=font(52, True), fill="#FFFFFF", anchor="mm")
    draw.text(
        (540, 1750),
        "Adaptável à sua empresa. Custos atrativos.",
        font=font(29),
        fill="#AEB8BF",
        anchor="mm",
    )
    return canvas


def localized_stock_slide(
    logo: Image.Image,
    materials: Path,
    purchasing: Path,
    language: str,
) -> Image.Image:
    if language == "es":
        kicker = "COMPRAS, MATERIA PRIMA Y STOCKS"
        title = "Nada falla. El stock acompaña."
        subtitle = "Necesidades, compras, entradas, reservas y valoración conectadas."
        band_title = "COMPRAR BIEN. ACTUALIZAR AUTOMÁTICAMENTE."
        band_text = "De la necesidad a la recepción, con el stock siempre actualizado."
        band_end = "Más control. Menos roturas."
    else:
        kicker = "COMPRAS, MATÉRIA-PRIMA E STOCKS"
        title = "Nada falha. O stock acompanha."
        subtitle = "Necessidades, compras, entradas, reservas e valorização ligadas."
        band_title = "COMPRAR CERTO. ATUALIZAR AUTOMATICAMENTE."
        band_text = "Da necessidade à receção, com stock sempre atual."
        band_end = "Mais controlo. Menos ruturas."

    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 72), kicker, font=font(25, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=LINE, width=2)
    draw.text((56, 245), title, font=font(52, True), fill=DARKER)
    draw.text((58, 330), subtitle, font=font(28), fill=MUTED)
    paste_card(canvas, purchasing, (56, 430, 1024, 930))
    paste_card(canvas, materials, (56, 970, 1024, 1470))
    draw.rounded_rectangle((56, 1515, 1024, 1760), 28, fill=GREEN)
    draw.text((540, 1587), band_title, font=font(30, True), fill=DARKER, anchor="mm")
    draw.text((540, 1652), band_text, font=font(27), fill=DARKER, anchor="mm")
    draw.text((540, 1710), band_end, font=font(28, True), fill=DARKER, anchor="mm")
    return canvas


def localized_finance_slide(
    logo: Image.Image,
    dashboard: Path,
    language: str,
) -> Image.Image:
    if language == "es":
        kicker = "CONTROL FINANCIERO"
        title = "¿Sabe cuánto vale su empresa?"
        subtitle = "Stock valorado, compras, ventas,\ncompromisos y márgenes en una sola visión."
        metrics = (("STOCK", "VALORADO"), ("COMPRAS", "CONTROLADAS"), ("MÁRGENES", "VISIBLES"))
        band_title = "CONTROL FINANCIERO PARA DECIDIR CON CONFIANZA."
        band_text = "Una lectura clara del presente. Una base sólida para crecer."
    else:
        kicker = "CONTROLO FINANCEIRO"
        title = "Sabe quanto vale a sua empresa?"
        subtitle = "Stock valorizado, compras, vendas,\ncompromissos e margens numa única visão."
        metrics = (("STOCK", "VALORIZADO"), ("COMPRAS", "CONTROLADAS"), ("MARGENS", "VISÍVEIS"))
        band_title = "CONTROLO FINANCEIRO PARA DECIDIR COM CONFIANÇA."
        band_text = "Uma leitura clara do presente. Uma base sólida para crescer."

    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    draw = ImageDraw.Draw(canvas)
    logo_mark(canvas, logo)
    draw.text((230, 72), kicker, font=font(27, True), fill=ORANGE)
    draw.line((56, 190, 1024, 190), fill=LINE, width=2)
    draw.text((56, 245), title, font=font(52, True), fill=DARKER)
    draw.multiline_text((58, 340), subtitle, font=font(30), fill=MUTED, spacing=10)
    paste_card(canvas, dashboard, (56, 535, 1024, 1315))
    for (label, value), x in zip(metrics, (56, 381, 706)):
        draw.rounded_rectangle((x, 1380, x + 294, 1545), 24, fill="#FFFFFF", outline=LINE, width=2)
        draw.text((x + 147, 1427), label, font=font(22, True), fill=ORANGE, anchor="mm")
        draw.text((x + 147, 1490), value, font=font(25, True), fill=DARKER, anchor="mm")
    draw.rounded_rectangle((56, 1590, 1024, 1760), 28, fill=DARKER)
    draw.text((540, 1648), band_title, font=font(28, True), fill="#FFFFFF", anchor="mm")
    draw.text((540, 1710), band_text, font=font(25), fill="#D8E0E5", anchor="mm")
    return canvas


def partner_closing_slide(
    wordmark: Image.Image,
    partner: Image.Image,
    language: str,
) -> Image.Image:
    """Fecho co-branded: produto em primeiro plano, representação oficial inequívoca."""
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    draw.rectangle((0, 0, WIDTH, 18), fill=ORANGE)
    for y in range(HEIGHT):
        alpha = int(52 * (y / HEIGHT))
        draw.line((0, y, WIDTH, y), fill=(23 + alpha // 5, 34 + alpha // 7, 45 + alpha // 8, 255))
    paste_wordmark(canvas, wordmark, (850, 300), (540, 245))
    draw.line((150, 435, 930, 435), fill=(255, 255, 255, 70), width=2)

    if language == "es":
        question = "¿POR QUÉ ELEGIR luGEST?"
        body = "Experiencia real de fábrica, tecnología industrial\ny acompañamiento cercano para transformar su gestión."
        partner_label = "REPRESENTANTE OFICIAL"
        cta = "SOLICITE UNA DEMOSTRACIÓN"
        footer = "ADAPTABLE A LA REALIDAD DE SU EMPRESA"
    else:
        question = "PORQUÊ ESCOLHER O luGEST?"
        body = "Experiência real de fábrica, tecnologia industrial\ne acompanhamento próximo para transformar a sua gestão."
        partner_label = "REPRESENTANTE OFICIAL"
        cta = "AGENDE UMA DEMONSTRAÇÃO"
        footer = "ADAPTÁVEL À REALIDADE DA SUA EMPRESA"

    draw.text((540, 555), question, font=font(45, True), fill="#FFFFFF", anchor="mm")
    draw.multiline_text(
        (540, 725), body, font=font(32), fill="#D9E0E4", anchor="mm", align="center", spacing=14
    )
    draw.text((540, 930), partner_label, font=font(24, True), fill="#FF6B62", anchor="mm")

    card = (92, 985, 988, 1375)
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle((106, 1003, 1002, 1393), 34, fill=(0, 0, 0, 85))
    canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(16)))
    draw.rounded_rectangle(card, 34, fill="#FFFFFF", outline="#D9DEDA", width=2)
    mark = partner.copy()
    mark.thumbnail((770, 280), Image.Resampling.LANCZOS)
    canvas.alpha_composite(mark, ((WIDTH - mark.width) // 2, 1040 + (270 - mark.height) // 2))

    draw.rounded_rectangle((125, 1470, 955, 1605), 67, fill=GREEN)
    draw.text((540, 1537), cta, font=font(31, True), fill=DARKER, anchor="mm")
    draw.text((540, 1740), footer, font=font(27, True), fill="#E9EDEE", anchor="mm")
    draw.text(
        (540, 1800),
        "luGEST × LASER IBERIC STUDIO SL",
        font=font(23, True),
        fill="#F0A169",
        anchor="mm",
    )
    return canvas


def closing_slide(logo: Image.Image) -> Image.Image:
    """Encerramento com uma pergunta forte e credibilidade industrial."""
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    mark = logo.copy()
    mark.thumbnail((300, 300), Image.Resampling.LANCZOS)
    canvas.alpha_composite(mark, ((WIDTH - mark.width) // 2, 120))
    draw.text((540, 520), "luGEST", font=font(100, True), fill="#FFFFFF", anchor="mm")
    draw.text((540, 610), "SOFTWARE ERP INDUSTRIAL", font=font(27, True), fill=ORANGE, anchor="mm")
    draw.line((160, 685, 920, 685), fill=(255, 255, 255, 72), width=2)
    draw.text(
        (540, 810),
        "PORQUÊ ESCOLHER O luGEST?",
        font=font(46, True),
        fill="#FFFFFF",
        anchor="mm",
    )
    draw.multiline_text(
        (540, 1060),
        "Porque nasceu da experiência real\nde quem conhece o chão de fábrica.",
        font=font(42, True),
        fill="#F5F6F3",
        anchor="mm",
        align="center",
        spacing=20,
    )
    draw.multiline_text(
        (540, 1285),
        "Conhecimento para organizar, controlar\ne fazer crescer a sua produção.",
        font=font(31),
        fill="#C9D1D7",
        anchor="mm",
        align="center",
        spacing=14,
    )
    draw.rounded_rectangle((105, 1450, 975, 1580), 65, fill=GREEN)
    draw.text(
        (540, 1515),
        "CONHECIMENTO DE FÁBRICA. CONTROLO PARA CRESCER.",
        font=font(25, True),
        fill=DARKER,
        anchor="mm",
    )
    draw.text(
        (540, 1730),
        "ADAPTÁVEL À REALIDADE DA SUA EMPRESA",
        font=font(27, True),
        fill="#E5EAED",
        anchor="mm",
    )
    draw.text(
        (540, 1790),
        "Agende uma demonstração.",
        font=font(25),
        fill="#AEB8BF",
        anchor="mm",
    )
    return canvas


def create_premium_music(path: Path, duration: float) -> None:
    """Cria uma cama musical limpa: apenas tons suaves, sem ruído ou percussão."""
    sample_rate = 48_000
    sample_count = int(duration * sample_rate)
    t = np.arange(sample_count, dtype=np.float64) / sample_rate
    left = np.zeros(sample_count, dtype=np.float64)
    right = np.zeros(sample_count, dtype=np.float64)

    # Progressão luminosa em Ré maior: D, A, Bm, G. Usamos apenas senoides
    # arredondadas; não há ruído, pratos, cliques ou transientes agressivos.
    progression = [
        (146.83, 185.00, 220.00),
        (110.00, 138.59, 164.81),
        (123.47, 146.83, 185.00),
        (98.00, 123.47, 146.83),
    ]
    block = 8.0
    blocks = math.ceil(duration / block)
    for block_index in range(blocks):
        start = block_index * block
        end = min(duration, start + block)
        chord = progression[block_index % len(progression)]
        active = (t >= start) & (t < end)
        local = np.clip(t - start, 0.0, block)
        attack = np.clip(local / 1.7, 0.0, 1.0)
        release = np.clip((end - t) / 1.8, 0.0, 1.0)
        env = active * np.minimum(attack, release)

        # Pad quente e largo, limitado às frequências médias/graves.
        for note_index, frequency in enumerate(chord):
            phase = note_index * 0.17
            left += 0.026 * env * np.sin(2 * np.pi * frequency * t + phase)
            right += 0.026 * env * np.sin(2 * np.pi * frequency * t + phase + 0.06)

        root = chord[0] / 2.0
        left += 0.020 * env * np.sin(2 * np.pi * root * t)
        right += 0.020 * env * np.sin(2 * np.pi * root * t + 0.03)

        # Motivo melódico lento, com ataque e saída longos para evitar estalidos.
        for step in range(4):
            note_start = start + 0.8 + step * 1.75
            if note_start >= end:
                break
            first = int(note_start * sample_rate)
            last = min(sample_count, first + int(1.45 * sample_rate))
            local_note = np.arange(last - first, dtype=np.float64) / sample_rate
            note_env = np.sin(np.pi * np.clip(local_note / 1.45, 0.0, 1.0)) ** 2
            frequency = chord[(step + 1) % len(chord)] * 2.0
            tone = np.sin(2 * np.pi * frequency * local_note)
            if step % 2:
                left[first:last] += 0.003 * note_env * tone
                right[first:last] += 0.007 * note_env * tone
            else:
                left[first:last] += 0.007 * note_env * tone
                right[first:last] += 0.003 * note_env * tone

    # Crescimento muito controlado e fades longos.
    build = 0.76 + 0.24 * np.clip(t / duration, 0.0, 1.0)
    finale = 1.0 + 0.12 * np.clip((t - (duration - 10.0)) / 8.0, 0.0, 1.0)
    fade = np.minimum(np.clip(t / 2.4, 0.0, 1.0), np.clip((duration - t) / 3.2, 0.0, 1.0))
    stereo = np.column_stack((left, right)) * (build * finale * fade)[:, None]
    peak = float(np.max(np.abs(stereo))) or 1.0
    # Normaliza a própria composição antes da mistura. Assim a música mantém
    # presença constante sem depender do nível bruto dos osciladores.
    stereo *= 0.38 / peak
    pcm = np.int16(np.clip(stereo, -1.0, 1.0) * 32767)
    with wave.open(str(path), "wb") as output:
        output.setnchannels(2)
        output.setsampwidth(2)
        output.setframerate(sample_rate)
        output.writeframes(pcm.tobytes())


async def create_voiceover(
    path: Path,
    text: str,
    voice: str,
    rate: str = "+11%",
) -> None:
    import edge_tts

    communication = edge_tts.Communicate(
        text=text,
        voice=voice,
        rate=rate,
        pitch="+0Hz",
        volume="+0%",
    )
    await communication.save(str(path))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ffmpeg", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--work", type=Path, required=True)
    parser.add_argument(
        "--music",
        type=Path,
        help="Faixa musical licenciada. Se omitida, é criada uma cama instrumental local.",
    )
    parser.add_argument("--language", choices=("pt", "es"), default="pt")
    parser.add_argument("--partner-logo", type=Path, required=True)
    args = parser.parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.work.mkdir(parents=True, exist_ok=True)

    temp = Path(os.environ.get("TEMP", "."))
    captures = {
        "materials": temp / "codex-clipboard-6b09b803-3e8b-4d01-899c-a45d2bbcfe8d.png",
        "products": temp / "codex-clipboard-f12a5c18-7cf3-4b03-b93f-0edbcf4929f4.png",
        "purchasing": temp / "codex-clipboard-a9d2dfe4-92cc-4c40-8cd9-684667008492.png",
        "quote": temp / "codex-clipboard-857a8b99-c634-4c67-90f4-67b6b79def9f.png",
        "nest": temp / "codex-clipboard-8ae4a5c5-bb42-4a64-84a3-108b81d661fe.png",
        "planning": temp / "codex-clipboard-2a19c2b6-f5e7-4e16-be59-72bfa41f28ff.png",
        "dashboard": temp / "codex-clipboard-ad5b90b8-e525-4d4b-9e11-79d0bb537962.png",
    }
    missing = [str(path) for path in captures.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("Capturas em falta:\n" + "\n".join(missing))

    logo_path = Path(__file__).resolve().parents[1] / "Logos" / "lugest_symbol.png"
    wordmark_path = Path(__file__).resolve().parents[1] / "Logos" / "lg.png"
    logo = Image.open(logo_path).convert("RGBA")
    wordmark = transparent_wordmark(wordmark_path)
    if not args.partner_logo.exists():
        raise FileNotFoundError(f"Logótipo do parceiro não encontrado: {args.partner_logo}")
    partner = transparent_partner_logo(args.partner_logo)

    if args.language == "es":
        copy = {
            "quote": (
                "Presupuestación industrial",
                "Presupuestos rápidos. Completos.",
                "DXF/DWG, operaciones, productos industriales, máquinas y conjuntos.",
                "PRESUPUESTAR RÁPIDO. DECIDIR CON RIGOR.",
            ),
            "nest": (
                "Nesting láser",
                "¿Aprovecha al máximo cada chapa?",
                "El nesting mide el consumo real de materia prima y lo lleva directamente al presupuesto.",
                "MENOS DESPERDICIO. MÁS MARGEN. MÁS CONFIANZA.",
            ),
            "planning": (
                "Planificación industrial",
                "Capacidad, prioridades y plazos.",
                "Máquinas, carga, bloqueos y semanas coordinados en una visión clara.",
                "LA FÁBRICA SABE QUÉ HACER, CUÁNDO Y DÓNDE.",
            ),
        }
        narration_blocks = [
            "Más producción, mejores decisiones y más control con luGEST.",
            (
                "Cree presupuestos industriales completos en minutos. Importe de equis efe y de uve doble ge. "
                "Calcule operaciones, productos, máquinas y conjuntos con precisión."
            ),
            (
                "El nésting calcula el consumo real de chapa y retales, reduce el desperdicio "
                "y convierte cada propuesta en un coste fiable."
            ),
            (
                "Planifique máquinas, capacidad, prioridades y plazos. "
                "Todo el equipo sabe qué hacer, cuándo y dónde."
            ),
            (
                "Compras, materia prima y stocks trabajan juntos. luGEST anticipa necesidades, "
                "evita faltas y mantiene las existencias actualizadas."
            ),
            (
                "En un único panel, controle ventas, costes, márgenes, producción y resultados. "
                "Más control. Más transparencia. Mejores decisiones."
            ),
            (
                "¿Por qué elegir luGEST? Porque combina la experiencia real de fábrica con tecnología industrial. "
                "Y con Laser Iberic Studio, representante oficial, ofrece un acompañamiento cercano, "
                "desde la demostración hasta la adaptación a su empresa. luGEST. Control para crecer."
            ),
        ]
        voice_name = "es-ES-ElviraNeural"
    else:
        copy = {
            "quote": (
                "Orçamentação industrial",
                "Orçamentos rápidos. Completos.",
                "DXF/DWG, operações, produtos industriais, máquinas e conjuntos.",
                "ORÇAMENTAR DEPRESSA. DECIDIR COM RIGOR.",
            ),
            "nest": (
                "Nesting Laser",
                "Está a aproveitar cada chapa?",
                "O nesting mede o consumo real de matéria-prima e leva-o diretamente ao orçamento.",
                "MENOS DESPERDÍCIO. MAIS MARGEM. MAIS CONFIANÇA.",
            ),
            "planning": (
                "Planeamento industrial",
                "Capacidade, prioridades e prazos.",
                "Máquinas, carga, bloqueios e semanas coordenados numa visão clara.",
                "A FÁBRICA SABE O QUE FAZER, QUANDO E ONDE.",
            ),
        }
        narration_blocks = [
            "Mais produção, melhores decisões e mais controlo com o luGEST.",
            (
                "Crie orçamentos industriais completos em minutos. Importe dê xis éfe e dê dâblio gê. "
                "Calcule operações, produtos, máquinas e conjuntos com rigor."
            ),
            (
                "O nésting calcula o consumo real de chapa e retalhos, reduz desperdício "
                "e transforma cada proposta num custo fiável."
            ),
            (
                "Planeie máquinas, capacidade, prioridades e prazos. "
                "Toda a equipa sabe o que fazer, quando e onde."
            ),
            (
                "Compras, matéria-prima e stocks trabalham em conjunto. O luGEST antecipa necessidades, "
                "evita faltas e mantém as existências atualizadas."
            ),
            (
                "Num único painel, acompanhe vendas, custos, margens, produção e resultados. "
                "Mais controlo. Mais transparência. Melhores decisões."
            ),
            (
                "Porquê escolher o luGEST? Porque combina a experiência real de fábrica com tecnologia industrial. "
                "E com a Laser Iberic Studio, representante oficial, oferece acompanhamento próximo, "
                "desde a demonstração até à adaptação à sua empresa. luGEST. Controlo para crescer."
            ),
        ]
        voice_name = "pt-PT-RaquelNeural"

    content_slides = [
        content_slide(logo, captures["quote"], *copy["quote"]),
        content_slide(logo, captures["nest"], *copy["nest"]),
        content_slide(logo, captures["planning"], *copy["planning"]),
        localized_stock_slide(logo, captures["materials"], captures["purchasing"], args.language),
        localized_finance_slide(logo, captures["dashboard"], args.language),
    ]
    slides = [clean_brand_intro_slide(wordmark)]
    slides.extend(add_partner_footer(slide, partner, args.language) for slide in content_slides)
    slides.append(partner_closing_slide(wordmark, partner, args.language))

    slide_paths: list[Path] = []
    for index, slide in enumerate(slides):
        path = args.work / f"short_slide_{index:02d}.png"
        slide.convert("RGB").save(path, quality=96)
        slide_paths.append(path)

    durations = [6.2, 11.5, 7.8, 8.2, 10.5, 12.8, 20.5]
    transition_duration = 0.5
    video_duration = sum(durations) - transition_duration * (len(durations) - 1)
    narration_starts: list[float] = []
    cursor = 0.0
    for index, duration in enumerate(durations):
        narration_starts.append(cursor + 0.35)
        cursor += duration - (transition_duration if index < len(durations) - 1 else 0.0)
    narration_windows = [duration - 1.0 for duration in durations]

    voice_paths: list[Path] = []
    for index, (text, safe_window) in enumerate(zip(narration_blocks, narration_windows)):
        raw_voice_path = args.work / f"short_narracao_{args.language}_{index:02d}_raw.mp3"
        voice_path = args.work / f"short_narracao_{args.language}_{index:02d}.wav"
        voice_rate = "+12%" if index == 1 else "+7%"
        asyncio.run(create_voiceover(raw_voice_path, text, voice=voice_name, rate=voice_rate))
        run(
            [
                str(args.ffmpeg),
                "-y",
                "-i",
                str(raw_voice_path),
                "-af",
                (
                    "silenceremove=start_periods=1:start_duration=0.04:start_threshold=-48dB,"
                    "areverse,"
                    "silenceremove=start_periods=1:start_duration=0.08:start_threshold=-48dB,"
                    "areverse,apad=pad_dur=0.12"
                ),
                "-ar",
                "48000",
                "-ac",
                "2",
                str(voice_path),
            ]
        )
        voice_duration = probe_duration(args.ffmpeg, voice_path)
        if voice_duration > safe_window:
            raise RuntimeError(
                f"O bloco de narração {index + 1} tem {voice_duration:.2f}s "
                f"e ultrapassa a janela segura de {safe_window:.2f}s."
            )
        voice_paths.append(voice_path)

    # Máscara móvel: revela os títulos da esquerda para a direita, como se
    # estivessem a ser carregados, sem sacrificar a legibilidade.
    title_curtain = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    ImageDraw.Draw(title_curtain).rectangle((36, 220, 1044, 510), fill=OFF_WHITE)
    title_curtain_path = args.work / "short_title_curtain.png"
    title_curtain.save(title_curtain_path)

    segment_paths: list[Path] = []
    for index, (slide_path, duration) in enumerate(zip(slide_paths, durations)):
        segment = args.work / f"short_segment_{index:02d}.mp4"
        if index == 0:
            render_brand_intro(
                args.ffmpeg,
                segment,
                args.work,
                wordmark,
                logo,
                duration,
                args.language,
            )
            segment_paths.append(segment)
            continue
        total_frames = max(1, round(duration * FPS))
        # Base rigorosamente estável. O movimento fica reservado às transições
        # e aos elementos gráficos, evitando vibração por reamostragem contínua.
        zoom_filter = (
            f"scale={WIDTH}:{HEIGHT}:flags=lanczos,"
            f"fps={FPS},format=yuv420p"
        )
        reveal_title = index in {1, 2, 3, 4, 5}
        if reveal_title:
            video_filter = (
                f"[0:v]{zoom_filter}[base];"
                "[1:v]format=rgba[curtain];"
                "[base][curtain]overlay="
                "x='min(1180,t*920)':y=0:shortest=1,format=yuv420p[out]"
            )
            input_args = [
                "-loop",
                "1",
                "-i",
                str(slide_path),
                "-loop",
                "1",
                "-i",
                str(title_curtain_path),
            ]
            filter_args = ["-filter_complex", video_filter, "-map", "[out]"]
        else:
            input_args = ["-loop", "1", "-i", str(slide_path)]
            filter_args = ["-vf", zoom_filter]
        run(
            [
                str(args.ffmpeg),
                "-y",
                *input_args,
                "-t",
                f"{duration:.2f}",
                *filter_args,
                "-r",
                str(FPS),
                "-c:v",
                "libx264",
                "-preset",
                "medium",
                "-crf",
                "18",
                str(segment),
            ]
        )
        segment_paths.append(segment)

    silent_video = args.work / "short_video_sem_audio.mp4"
    transition_offsets: list[float] = []
    transition_cursor = durations[0] - transition_duration
    for duration in durations[1:]:
        transition_offsets.append(transition_cursor)
        transition_cursor += duration - transition_duration
    transitions = [
        "fadeblack",
        "smoothleft",
        "smoothup",
        "diagtl",
        "smoothright",
        "fadeblack",
    ]
    transition_chain: list[str] = []
    previous = "[0:v]"
    for index, (offset, transition) in enumerate(zip(transition_offsets, transitions), start=1):
        output_label = f"[v{index}]"
        transition_chain.append(
            f"{previous}[{index}:v]xfade=transition={transition}:"
            f"duration={transition_duration}:offset={offset}{output_label}"
        )
        previous = output_label
    run(
        [
            str(args.ffmpeg),
            "-y",
            *sum((["-i", str(path)] for path in segment_paths), []),
            "-filter_complex",
            ";".join(transition_chain),
            "-map",
            previous,
            "-t",
            f"{video_duration:.2f}",
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-crf",
            "18",
            "-pix_fmt",
            "yuv420p",
            str(silent_video),
        ]
    )

    if args.music:
        if not args.music.exists():
            raise FileNotFoundError(f"Faixa musical não encontrada: {args.music}")
        music_path = args.music
    else:
        music_path = args.work / "short_musica_premium_final.wav"
        create_premium_music(music_path, video_duration)
    audio_inputs: list[str] = []
    for path in voice_paths:
        audio_inputs.extend(["-i", str(path)])
    delayed_voices: list[str] = []
    filter_parts: list[str] = []
    for index, start in enumerate(narration_starts, start=1):
        delay_ms = round(start * 1000)
        label = f"voice{index}"
        filter_parts.append(
            f"[{index}:a]volume=1.02,highpass=f=78,lowpass=f=12500,"
            "acompressor=threshold=0.16:ratio=2.0:attack=16:release=180:makeup=1.08,"
            f"aformat=sample_rates=48000:channel_layouts=stereo,"
            f"adelay={delay_ms}|{delay_ms}[{label}]"
        )
        delayed_voices.append(f"[{label}]")
    music_input_index = len(voice_paths) + 1
    filter_parts.append(
        f"[{music_input_index}:a]atrim=0:{video_duration:.2f},volume=0.18,"
        f"afade=t=in:st=0:d=1.3,afade=t=out:st={video_duration - 3.0:.2f}:d=3.0,"
        "aformat=sample_rates=48000:sample_fmts=fltp:channel_layouts=stereo[music]"
    )
    filter_parts.append(
        "".join(delayed_voices)
        + f"amix=inputs={len(delayed_voices)}:duration=longest:normalize=0,"
        "aformat=sample_rates=48000:sample_fmts=fltp:channel_layouts=stereo[voiceout]"
    )
    filter_parts.append(
        "[voiceout][music]amix=inputs=2:duration=longest:dropout_transition=0:normalize=0,"
        "alimiter=limit=0.93:attack=6:release=80,"
        "aformat=sample_rates=48000:sample_fmts=fltp:channel_layouts=stereo[a]"
    )
    run(
        [
            str(args.ffmpeg),
            "-y",
            "-i",
            str(silent_video),
            *audio_inputs,
            "-i",
            str(music_path),
            "-filter_complex",
            ";".join(filter_parts),
            "-map",
            "0:v:0",
            "-map",
            "[a]",
            "-t",
            f"{video_duration:.2f}",
            "-c:v",
            "copy",
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-movflags",
            "+faststart",
            str(args.output),
        ]
    )
    print(args.output)


if __name__ == "__main__":
    main()
