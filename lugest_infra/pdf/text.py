from __future__ import annotations

from typing import Any


def clip_text(value: Any, max_width: float, font_name: str, font_size: float) -> str:
    from reportlab.pdfbase import pdfmetrics

    text = "" if value is None else str(value)
    if pdfmetrics.stringWidth(text, font_name, font_size) <= max_width:
        return text
    ellipsis = "..."
    while text and pdfmetrics.stringWidth(text + ellipsis, font_name, font_size) > max_width:
        text = text[:-1]
    return f"{text}{ellipsis}" if text else ""


def wrap_text(value: Any, font_name: str, font_size: float, max_width: float, max_lines: int | None = None) -> list[str]:
    from reportlab.pdfbase import pdfmetrics

    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    if not text:
        return []
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip() if current else word
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
        if max_lines and len(lines) >= max_lines:
            break
    if current and (not max_lines or len(lines) < max_lines):
        lines.append(current)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
    return lines


def fit_font_size(text: Any, font_name: str, max_width: float, preferred_size: float, min_size: float) -> float:
    from reportlab.pdfbase import pdfmetrics

    size = float(preferred_size)
    raw = str(text or "")
    max_width = max(12.0, float(max_width))
    while size > float(min_size) and pdfmetrics.stringWidth(raw, font_name, size) > max_width:
        size -= 0.3
    return max(float(min_size), round(size, 2))

