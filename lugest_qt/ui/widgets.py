from __future__ import annotations

import re

from PySide6.QtGui import QValidator
from PySide6.QtWidgets import QDoubleSpinBox, QFrame, QLabel, QVBoxLayout


class FlexibleDecimalSpinBox(QDoubleSpinBox):
    """QDoubleSpinBox that accepts both comma and dot as decimal separators."""

    def _strip_affixes(self, text: str) -> str:
        clean = str(text or "").replace("\u00a0", " ").strip()
        prefix = str(self.prefix() or "")
        suffix = str(self.suffix() or "")
        if prefix and clean.lower().startswith(prefix.lower()):
            clean = clean[len(prefix) :].strip()
        if suffix and clean.lower().endswith(suffix.lower()):
            clean = clean[: -len(suffix)].strip()
        clean = re.sub(r"(?i)\b(eur|euro|euros)\b", "", clean)
        clean = clean.replace("\u20ac", "").strip()
        return clean

    def _normalise_decimal_text(self, text: str) -> str:
        clean = self._strip_affixes(text).replace(" ", "")
        if not clean:
            return ""
        match = re.search(r"[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)(?:[.,]\d+)*", clean)
        if not match:
            return clean
        clean = match.group(0)
        if "," in clean and "." in clean:
            # Treat the rightmost separator as decimal and the other as thousands.
            if clean.rfind(",") > clean.rfind("."):
                clean = clean.replace(".", "").replace(",", ".")
            else:
                clean = clean.replace(",", "")
        elif clean.count(",") > 1:
            head, tail = clean.rsplit(",", 1)
            clean = head.replace(",", "") + "." + tail
        elif clean.count(".") > 1:
            head, tail = clean.rsplit(".", 1)
            clean = head.replace(".", "") + "." + tail
        else:
            clean = clean.replace(",", ".")
        return clean

    def validate(self, text: str, pos: int) -> tuple[QValidator.State, str, int]:
        normalized = self._normalise_decimal_text(text)
        if normalized in {"", "-", "+", ".", "-.", "+."}:
            return (QValidator.Intermediate, text, pos)
        try:
            value = float(normalized)
        except ValueError:
            return (QValidator.Invalid, text, pos)
        if self.minimum() <= value <= self.maximum():
            return (QValidator.Acceptable, text, pos)
        return (QValidator.Intermediate, text, pos)

    def valueFromText(self, text: str) -> float:
        normalized = self._normalise_decimal_text(text)
        try:
            return float(normalized)
        except ValueError:
            return float(self.value())


class CardFrame(QFrame):
    def __init__(self, parent=None, object_name: str = "Card") -> None:
        super().__init__(parent)
        self.setObjectName(object_name)
        self.setProperty("tone", "default")

    def set_tone(self, tone: str = "default") -> None:
        self.setProperty("tone", str(tone or "default"))
        style = self.style()
        if style is not None:
            style.unpolish(self)
            style.polish(self)
        self.update()


class StatCard(CardFrame):
    def __init__(self, title: str, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(4)
        self.accent = QFrame()
        self.accent.setFixedHeight(3)
        self.accent.setStyleSheet("background: #c6d2e0; border-radius: 1px;")
        self.title_label = QLabel(title)
        self.title_label.setProperty("role", "muted")
        self.title_label.setStyleSheet("font-size: 10px; letter-spacing: 0.2px;")
        self.value_label = QLabel("-")
        self.value_label.setStyleSheet("font-size: 18px; font-weight: 700; color: #0f172a;")
        self.subtitle_label = QLabel("")
        self.subtitle_label.setProperty("role", "muted")
        self.subtitle_label.setStyleSheet("font-size: 10px;")
        self.subtitle_label.setWordWrap(False)
        layout.addWidget(self.accent)
        layout.addWidget(self.title_label)
        layout.addWidget(self.value_label)
        layout.addWidget(self.subtitle_label)
        layout.addStretch(1)
        self.set_tone("default")

    def set_data(self, value: str, subtitle: str = "") -> None:
        self.value_label.setText(str(value))
        self.subtitle_label.setText(str(subtitle))

    def set_tone(self, tone: str = "default") -> None:
        super().set_tone(tone)
        tone_map = {
            "default": ("#b8bcb8", "#30343b"),
            "info": ("#6f7771", "#30343b"),
            "success": ("#70c51a", "#30343b"),
            "warning": ("#d18a1f", "#30343b"),
            "danger": ("#555955", "#30343b"),
            "rejected": ("#d92d20", "#30343b"),
        }
        accent, value_color = tone_map.get(str(tone or "default"), tone_map["default"])
        self.accent.setStyleSheet(f"background: {accent}; border-radius: 2px;")
        self.value_label.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {value_color};")
