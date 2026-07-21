from __future__ import annotations

from lugest_infra.pdf.font_policy import install_pdf_font_policy

install_pdf_font_policy()

from .main_bridge import LegacyBackend

__all__ = ["LegacyBackend"]
