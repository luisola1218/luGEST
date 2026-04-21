from __future__ import annotations

from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout


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
            "default": ("#c6d2e0", "#0f172a"),
            "info": ("#6b8fb3", "#16324b"),
            "success": ("#4b8f69", "#0f172a"),
            "warning": ("#b58a3c", "#0f172a"),
            "danger": ("#b8574d", "#0f172a"),
        }
        accent, value_color = tone_map.get(str(tone or "default"), tone_map["default"])
        self.accent.setStyleSheet(f"background: {accent}; border-radius: 2px;")
        self.value_label.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {value_color};")
