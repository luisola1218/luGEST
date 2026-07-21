from __future__ import annotations

from functools import wraps
from typing import Any, Callable


MAX_PDF_FONT_SIZE = 8.0


def _bounded_size(value: Any) -> Any:
    try:
        return min(float(value), MAX_PDF_FONT_SIZE)
    except (TypeError, ValueError):
        return value


def _install_on(cls: type, method_name: str) -> None:
    original: Callable[..., Any] = getattr(cls, method_name)
    if getattr(original, "_lugest_font_policy", False):
        return

    @wraps(original)
    def bounded(self, font_name, font_size, *args, **kwargs):
        return original(self, font_name, _bounded_size(font_size), *args, **kwargs)

    bounded._lugest_font_policy = True  # type: ignore[attr-defined]
    setattr(cls, method_name, bounded)


def _install_canvas_initial_size(canvas_cls: type) -> None:
    original: Callable[..., Any] = canvas_cls.__init__
    if getattr(original, "_lugest_font_policy", False):
        return

    @wraps(original)
    def bounded_init(self, *args, **kwargs):
        values = list(args)
        # Canvas positional index 11 is initialFontSize (excluding self).
        if len(values) > 11:
            values[11] = _bounded_size(values[11] if values[11] is not None else MAX_PDF_FONT_SIZE)
        else:
            initial_size = kwargs.get("initialFontSize")
            kwargs["initialFontSize"] = _bounded_size(initial_size if initial_size is not None else MAX_PDF_FONT_SIZE)
        return original(self, *values, **kwargs)

    bounded_init._lugest_font_policy = True  # type: ignore[attr-defined]
    canvas_cls.__init__ = bounded_init


def _install_canvas_preamble(canvas_cls: type) -> None:
    original: Callable[..., Any] = canvas_cls._make_preamble
    if getattr(original, "_lugest_font_policy", False):
        return

    @wraps(original)
    def bounded_preamble(self, *args, **kwargs):
        result = original(self, *args, **kwargs)
        # ReportLab hardcodes an unused 12 pt font in every page preamble.
        self._preamble = str(self._preamble).replace(" 12 Tf 14.4 TL", " 8 Tf 9.6 TL")
        return result

    bounded_preamble._lugest_font_policy = True  # type: ignore[attr-defined]
    canvas_cls._make_preamble = bounded_preamble


def install_pdf_font_policy() -> None:
    """Apply the document-wide 8 pt ceiling to every ReportLab text path."""
    from reportlab.pdfgen.canvas import Canvas
    from reportlab.pdfgen.textobject import PDFTextObject

    _install_canvas_initial_size(Canvas)
    _install_canvas_preamble(Canvas)
    _install_on(Canvas, "setFont")
    _install_on(PDFTextObject, "setFont")
