from __future__ import annotations

import argparse
import asyncio
import math
import os
import subprocess
import sys
import wave
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont


WIDTH = 1920
HEIGHT = 1080
FPS = 30

DARK = "#2F3B46"
DARKER = "#17222D"
ORANGE = "#F4690A"
GREEN = "#69C80E"
OFF_WHITE = "#F5F6F3"
MUTED = "#637181"
LINE = "#D8DDD9"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        Path(r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf"),
        Path(r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


F_TITLE = font(58, True)
F_SUBTITLE = font(30)
F_KICKER = font(20, True)
F_CAPTION = font(28)
F_SMALL = font(20)


def rounded_mask(size: tuple[int, int], radius: int) -> Image.Image:
    mask = Image.new("L", size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius, fill=255)
    return mask


def fit_image(path: Path, size: tuple[int, int], crop: bool = False) -> Image.Image:
    image = Image.open(path).convert("RGB")
    target_ratio = size[0] / size[1]
    ratio = image.width / image.height
    if crop:
        if ratio > target_ratio:
            new_width = round(image.height * target_ratio)
            left = (image.width - new_width) // 2
            image = image.crop((left, 0, left + new_width, image.height))
        else:
            new_height = round(image.width / target_ratio)
            top = (image.height - new_height) // 2
            image = image.crop((0, top, image.width, top + new_height))
        return image.resize(size, Image.Resampling.LANCZOS)
    image.thumbnail(size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", size, "#FFFFFF")
    canvas.paste(image, ((size[0] - image.width) // 2, (size[1] - image.height) // 2))
    return canvas


def paste_card(
    canvas: Image.Image,
    screenshot: Path,
    box: tuple[int, int, int, int],
    radius: int = 22,
    crop: bool = False,
) -> None:
    x1, y1, x2, y2 = box
    card_w, card_h = x2 - x1, y2 - y1
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((x1 + 14, y1 + 18, x2 + 14, y2 + 18), radius, fill=(23, 34, 45, 42))
    shadow = shadow.filter(ImageFilter.GaussianBlur(13))
    canvas.alpha_composite(shadow)

    content = fit_image(screenshot, (card_w, card_h), crop=crop).convert("RGBA")
    content.putalpha(rounded_mask((card_w, card_h), radius))
    canvas.alpha_composite(content, (x1, y1))
    ImageDraw.Draw(canvas).rounded_rectangle(box, radius, outline="#C8CFCA", width=2)


def header(canvas: Image.Image, logo: Image.Image, kicker: str, title: str, subtitle: str) -> None:
    draw = ImageDraw.Draw(canvas)
    logo_copy = logo.copy()
    logo_copy.thumbnail((170, 104), Image.Resampling.LANCZOS)
    canvas.alpha_composite(logo_copy, (70, 50))
    draw.text((270, 52), kicker.upper(), font=F_KICKER, fill=ORANGE)
    draw.text((270, 78), title, font=F_TITLE, fill=DARKER)
    draw.text((273, 151), subtitle, font=F_SUBTITLE, fill=MUTED)
    draw.line((70, 212, 1850, 212), fill=LINE, width=2)


def standard_slide(
    logo: Image.Image,
    screenshot: Path,
    kicker: str,
    title: str,
    subtitle: str,
    caption: str,
) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    header(canvas, logo, kicker, title, subtitle)
    paste_card(canvas, screenshot, (70, 245, 1850, 915), crop=False)
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((90, 875, 1830, 985), 18, fill=(23, 34, 45, 226))
    draw.rectangle((90, 875, 101, 985), fill=ORANGE)
    draw.text((130, 901), caption, font=F_CAPTION, fill="#FFFFFF")
    draw.text((72, 1025), "LUGEST  |  SOFTWARE ERP INDUSTRIAL", font=F_SMALL, fill=MUTED)
    return canvas


def split_slide(
    logo: Image.Image,
    left: Path,
    right: Path,
    kicker: str,
    title: str,
    subtitle: str,
    caption: str,
) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), OFF_WHITE)
    header(canvas, logo, kicker, title, subtitle)
    paste_card(canvas, left, (70, 260, 945, 855), crop=False)
    paste_card(canvas, right, (975, 260, 1850, 855), crop=False)
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((230, 885, 1690, 978), 18, fill=DARK)
    bbox = draw.textbbox((0, 0), caption, font=F_CAPTION)
    draw.text(((WIDTH - (bbox[2] - bbox[0])) / 2, 912), caption, font=F_CAPTION, fill="#FFFFFF")
    draw.text((72, 1025), "LUGEST  |  SOFTWARE ERP INDUSTRIAL", font=F_SMALL, fill=MUTED)
    return canvas


def cover_slide(logo: Image.Image) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    for index in range(16):
        alpha = max(0, 75 - index * 4)
        draw.ellipse(
            (1120 - index * 35, -180 - index * 15, 2140 + index * 35, 820 + index * 15),
            outline=(244, 105, 10, alpha),
            width=3,
        )
    logo_copy = logo.copy()
    logo_copy.thumbnail((420, 420), Image.Resampling.LANCZOS)
    canvas.alpha_composite(logo_copy, (110, 292))
    draw.text((620, 260), "LUGEST", font=font(100, True), fill="#FFFFFF")
    draw.text((625, 382), "SOFTWARE ERP INDUSTRIAL", font=font(34, True), fill=ORANGE)
    draw.text((625, 500), "O controlo começa aqui.", font=font(54, True), fill="#FFFFFF")
    draw.text((625, 575), "Transparência para decidir. Orientação para crescer.", font=font(31), fill="#C9D0D5")
    draw.rounded_rectangle((625, 700, 1190, 772), 34, fill=GREEN)
    draw.text((681, 716), "CONTROLO  •  TRANSPARÊNCIA  •  RESULTADOS", font=font(20, True), fill=DARKER)
    return canvas


def closing_slide(logo: Image.Image) -> Image.Image:
    canvas = Image.new("RGBA", (WIDTH, HEIGHT), DARKER)
    draw = ImageDraw.Draw(canvas)
    logo_copy = logo.copy()
    logo_copy.thumbnail((360, 360), Image.Resampling.LANCZOS)
    canvas.alpha_composite(logo_copy, ((WIDTH - logo_copy.width) // 2, 120))
    title = "Mais controlo. Mais transparência.\nMais resultados."
    draw.multiline_text((WIDTH // 2, 500), title, font=font(62, True), fill="#FFFFFF", anchor="mm", align="center", spacing=12)
    draw.rounded_rectangle((620, 700, 1300, 790), 45, fill=ORANGE)
    draw.text((960, 744), "COMECE HOJE A TRANSFORMAR A SUA EMPRESA", font=font(24, True), fill="#FFFFFF", anchor="mm")
    draw.text((960, 880), "luGEST  |  Software ERP Industrial", font=font(26), fill="#BBC4CA", anchor="mm")
    return canvas


def create_music(path: Path, duration: float) -> None:
    sample_rate = 48_000
    t = np.arange(int(duration * sample_rate), dtype=np.float64) / sample_rate
    left = np.zeros_like(t)
    right = np.zeros_like(t)

    # Uplifting D-major progression: D, B minor, G and A. Long pads provide
    # a cinematic corporate foundation without masking the narration.
    progression = [
        (146.83, 185.00, 220.00),
        (123.47, 146.83, 185.00),
        (98.00, 123.47, 146.83),
        (110.00, 138.59, 164.81),
    ]
    block = 8.0
    for i, chord in enumerate(progression * (math.ceil(duration / (block * len(progression))) + 1)):
        start = i * block
        if start >= duration:
            break
        envelope = np.clip((t - start) / 1.8, 0, 1) * np.clip((start + block - t) / 1.8, 0, 1)
        envelope[(t < start) | (t >= start + block)] = 0
        for frequency in chord:
            pad = envelope * (
                np.sin(2 * np.pi * frequency * t)
                + 0.20 * np.sin(2 * np.pi * frequency * 2 * t)
            )
            left += 0.013 * pad
            right += 0.013 * envelope * (
                np.sin(2 * np.pi * frequency * t + 0.12)
                + 0.20 * np.sin(2 * np.pi * frequency * 2 * t + 0.08)
            )

        # A warm bass root gives the presentation weight and forward motion.
        root = chord[0] / 2
        left += 0.014 * envelope * np.sin(2 * np.pi * root * t)
        right += 0.014 * envelope * np.sin(2 * np.pi * root * t + 0.05)

        # A restrained, alternating arpeggio adds optimism and momentum.
        for step in range(16):
            note_start = start + step * 0.5
            if note_start >= min(start + block, duration):
                break
            note_start_index = max(0, int(note_start * sample_rate))
            note_end_index = min(len(t), note_start_index + int(0.46 * sample_rate))
            local = np.arange(note_end_index - note_start_index, dtype=np.float64) / sample_rate
            note_env = np.exp(-local * 3.6) * np.clip(local / 0.018, 0, 1)
            note = chord[step % len(chord)] * 2
            tone = (
                np.sin(2 * np.pi * note * local)
                + 0.18 * np.sin(2 * np.pi * note * 2 * local)
            )
            if step % 2:
                right[note_start_index:note_end_index] += 0.010 * note_env * tone
                left[note_start_index:note_end_index] += 0.004 * note_env * tone
            else:
                left[note_start_index:note_end_index] += 0.010 * note_env * tone
                right[note_start_index:note_end_index] += 0.004 * note_env * tone

    # A soft heartbeat every two seconds makes the bed feel purposeful. The
    # envelope starts at zero, so it cannot introduce the clicks heard before.
    for beat_start in np.arange(0.0, duration, 2.0):
        beat_start_index = max(0, int(beat_start * sample_rate))
        beat_end_index = min(len(t), beat_start_index + int(0.42 * sample_rate))
        local = np.arange(beat_end_index - beat_start_index, dtype=np.float64) / sample_rate
        beat_env = np.exp(-local * 9.0) * np.clip(local / 0.012, 0, 1)
        beat = np.sin(2 * np.pi * (58.0 - 12.0 * np.clip(local, 0, 0.42)) * local)
        left[beat_start_index:beat_end_index] += 0.018 * beat_env * beat
        right[beat_start_index:beat_end_index] += 0.018 * beat_env * beat

    # The music grows gently towards the final call to action.
    build = 0.72 + 0.28 * np.clip(t / max(duration - 14.0, 1.0), 0, 1)
    finale = 1.0 + 0.22 * np.clip((t - (duration - 16.0)) / 12.0, 0, 1)
    fade = np.minimum(np.clip(t / 2.5, 0, 1), np.clip((duration - t) / 3.5, 0, 1))
    left *= build * finale * fade
    right *= build * finale * fade
    stereo = np.column_stack((left, right))
    peak = float(np.max(np.abs(stereo))) or 1.0
    stereo *= min(1.0, 0.82 / peak)
    pcm = np.int16(np.clip(stereo, -1, 1) * 32767)
    with wave.open(str(path), "wb") as output:
        output.setnchannels(2)
        output.setsampwidth(2)
        output.setframerate(sample_rate)
        output.writeframes(pcm.tobytes())


async def create_voiceover(path: Path, text: str) -> None:
    import edge_tts

    communication = edge_tts.Communicate(
        text=text,
        voice="pt-PT-RaquelNeural",
        rate="+8%",
        pitch="+4Hz",
        volume="+5%",
    )
    await communication.save(str(path))


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


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ffmpeg", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--work", type=Path, required=True)
    args = parser.parse_args()

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.work.mkdir(parents=True, exist_ok=True)

    temp = Path(os.environ.get("TEMP", "."))
    captures = {
        "dashboard": temp / "codex-clipboard-ad5b90b8-e525-4d4b-9e11-79d0bb537962.png",
        "materials": temp / "codex-clipboard-6b09b803-3e8b-4d01-899c-a45d2bbcfe8d.png",
        "products": temp / "codex-clipboard-f12a5c18-7cf3-4b03-b93f-0edbcf4929f4.png",
        "quote": temp / "codex-clipboard-857a8b99-c634-4c67-90f4-67b6b79def9f.png",
        "nest_setup": temp / "codex-clipboard-e7aa6dfa-3357-4ebe-9259-2223d6659e1b.png",
        "nest_result": temp / "codex-clipboard-8ae4a5c5-bb42-4a64-84a3-108b81d661fe.png",
        "planning": temp / "codex-clipboard-2a19c2b6-f5e7-4e16-be59-72bfa41f28ff.png",
        "transport": temp / "codex-clipboard-7da3c4e4-48e9-4780-8de5-1a9fe9837b06.png",
        "assistant": temp / "codex-clipboard-e677928a-44a6-41a7-973d-2fab6fc53252.png",
        "operator": temp / "codex-clipboard-0c2a490a-947d-417b-a38d-db0f2d2c2287.png",
        "pulse": temp / "codex-clipboard-0e39ca34-899a-47df-a60b-b83d7f17b0b8.png",
    }
    missing = [str(path) for path in captures.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("Capturas em falta:\n" + "\n".join(missing))

    logo_path = Path(__file__).resolve().parents[1] / "Logos" / "lugest_symbol.png"
    logo = Image.open(logo_path).convert("RGBA")

    slides: list[Image.Image] = [
        cover_slide(logo),
        standard_slide(
            logo,
            captures["dashboard"],
            "Decisão",
            "Uma visão única da empresa",
            "Património, vendas, compras e execução operacional em tempo real.",
            "Indicadores claros para decidir antes do problema chegar à produção.",
        ),
        split_slide(
            logo,
            captures["materials"],
            captures["products"],
            "Rastreabilidade",
            "Stock técnico, disponível e valorizado",
            "Matéria-prima e produto acabado ligados à operação.",
            "Lotes, reservas, preços, disponibilidade e histórico.",
        ),
        standard_slide(
            logo,
            captures["quote"],
            "Orçamentação",
            "Do DXF à proposta comercial",
            "Geometria, material, operações, tempos e custos na mesma linha.",
            "Orçamentos industriais verificáveis e preparados para produzir.",
        ),
        split_slide(
            logo,
            captures["nest_setup"],
            captures["nest_result"],
            "Nesting Laser",
            "Mais peças. Menos desperdício.",
            "Parâmetros, stock e otimização integrados no orçamento.",
            "Exploração de layouts com leitura clara da eficiência real.",
        ),
        standard_slide(
            logo,
            captures["planning"],
            "Planeamento",
            "Capacidade visível por máquina e semana",
            "Carga, bloqueios, prazos e prioridades num quadro operacional.",
            "O planeamento deixa de estar disperso e passa a orientar a fábrica.",
        ),
        standard_slide(
            logo,
            captures["assistant"],
            "Assistente MP",
            "Material certo, no momento certo",
            "Alertas de stock e recomendações ligadas ao horizonte produtivo.",
            "Menos ruturas, menos urgências e decisões de separação auditáveis.",
        ),
        standard_slide(
            logo,
            captures["operator"],
            "Chão de fábrica",
            "Execução simples para o operador",
            "Operações, progresso, consumos e avarias atualizados na origem.",
            "O estado real da produção chega à gestão sem folhas paralelas.",
        ),
        standard_slide(
            logo,
            captures["transport"],
            "Logística",
            "Rotas e entregas ligadas às encomendas",
            "Viagens próprias ou subcontratadas com destinos e custos controlados.",
            "Planeie, acompanhe e confirme cada entrega num fluxo único.",
        ),
        standard_slide(
            logo,
            captures["pulse"],
            "Desempenho",
            "O pulso da operação industrial",
            "OEE, disponibilidade, perdas, desvios e causas de paragem.",
            "Dados operacionais transformados em prioridades concretas.",
        ),
        closing_slide(logo),
    ]

    slide_paths: list[Path] = []
    for index, slide in enumerate(slides):
        path = args.work / f"slide_{index:02d}.png"
        slide.convert("RGB").save(path, quality=96)
        slide_paths.append(path)

    narration = (
        "O futuro da sua empresa começa com controlo! Com informação clara. Com decisões seguras. E com resultados! "
        "O luGEST é uma plataforma industrial adaptável à realidade de cada cliente, criada para ligar toda a operação: "
        "da proposta comercial, até à produção e à entrega. "
        "Tenha uma visão executiva transparente do património, das vendas, das compras e da operação, sempre atualizada. "
        "Gira clientes, fornecedores, produtos, encomendas e notas de compra num fluxo comum. "
        "Controle matéria-prima e produto acabado, com lotes, reservas, movimentos, preços e valorização rastreáveis. "
        "Na orçamentação, importe ficheiros D X F, valide a geometria e calcule material, operações, tempos, peso e custos. "
        "No Nesting Laser, explore layouts, aproveite stock e retalhos e reduza o desperdício de chapa. "
        "Planeie capacidade por máquina, semana e prioridade, com carga, bloqueios e prazos sempre visíveis. "
        "O Assistente de Matéria-Prima antecipa ruturas e ajuda a preparar o material certo, no momento certo. "
        "No chão de fábrica, o operador regista início e fim de operações, progresso, consumos, baixas e avarias diretamente na origem. "
        "Na qualidade, acompanhe ocorrências e mantenha a informação ligada à encomenda. "
        "Na logística, planeie viagens, custos, cada róta de transporte e a confirmação de cada entrega. "
        "E com o Pulse, acompanhe O E E, disponibilidade, perdas, desvios e causas de paragem. "
        "Dashboards, alertas e relatórios transformam dados operacionais em ações concretas de melhoria. "
        "Mais do que implementar software, é começar uma nova forma de controlar a sua empresa: "
        "mais simples, mais transparente e preparada para crescer consigo. "
        "luGEST. Controlo para decidir. Transparência para confiar. Orientação para resultados! "
        "Comece hoje. Agende uma demonstração!"
    )
    voice_path = args.work / "narracao_pt.mp3"
    asyncio.run(create_voiceover(voice_path, narration))
    voice_duration = probe_duration(args.ffmpeg, voice_path)
    total_duration = max(72.0, voice_duration + 3.0)
    per_slide = total_duration / len(slide_paths)

    segment_paths: list[Path] = []
    for index, slide_path in enumerate(slide_paths):
        segment = args.work / f"segment_{index:02d}.mp4"
        fade_out = max(0.1, per_slide - 0.55)
        run(
            [
                str(args.ffmpeg),
                "-y",
                "-loop",
                "1",
                "-i",
                str(slide_path),
                "-t",
                f"{per_slide:.3f}",
                "-vf",
                (
                    f"scale={WIDTH}:{HEIGHT}:flags=lanczos,"
                    f"fps={FPS},"
                    f"fade=t=in:st=0:d=0.40,fade=t=out:st={fade_out:.3f}:d=0.50,"
                    "format=yuv420p"
                ),
                "-an",
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

    concat_file = args.work / "segments.txt"
    concat_file.write_text(
        "\n".join(f"file '{str(path).replace(chr(39), chr(39) * 2)}'" for path in segment_paths),
        encoding="utf-8",
    )
    silent_video = args.work / "video_sem_audio.mp4"
    run(
        [
            str(args.ffmpeg),
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            str(concat_file),
            "-c",
            "copy",
            str(silent_video),
        ]
    )

    music_path = args.work / "musica_original.wav"
    create_music(music_path, total_duration)
    run(
        [
            str(args.ffmpeg),
            "-y",
            "-i",
            str(silent_video),
            "-i",
            str(voice_path),
            "-i",
            str(music_path),
            "-filter_complex",
            (
                "[1:a]volume=1.0,highpass=f=80,lowpass=f=11000,"
                "aformat=sample_rates=48000:channel_layouts=stereo[voice];"
                "[2:a]volume=0.75,highpass=f=42,lowpass=f=8500[music];"
                "[voice][music]amix=inputs=2:duration=longest:dropout_transition=2,"
                "loudnorm=I=-16:TP=-1.5:LRA=7,aresample=48000[a]"
            ),
            "-map",
            "0:v:0",
            "-map",
            "[a]",
            "-c:v",
            "copy",
            "-c:a",
            "aac",
            "-b:a",
            "192k",
            "-shortest",
            "-movflags",
            "+faststart",
            str(args.output),
        ]
    )
    print(args.output)


if __name__ == "__main__":
    main()
