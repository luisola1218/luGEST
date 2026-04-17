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
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(3)
        self.accent = QFrame()
        self.accent.setFixedHeight(4)
        self.accent.setStyleSheet("background: #c6d2e0; border-radius: 2px;")
        self.title_label = QLabel(title)
        self.title_label.setProperty("role", "muted")
        self.title_label.setStyleSheet("font-size: 11px;")
        self.value_label = QLabel("-")
        self.value_label.setStyleSheet("font-size: 20px; font-weight: 800; color: #0f172a;")
        self.subtitle_label = QLabel("")
        self.subtitle_label.setProperty("role", "muted")
        self.subtitle_label.setStyleSheet("font-size: 11px;")
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
            "info": ("#2f6db2", "#12304a"),
            "success": ("#1f8a4c", "#0f172a"),
            "warning": ("#c28112", "#0f172a"),
            "danger": ("#c24136", "#0f172a"),
        }
        accent, value_color = tone_map.get(str(tone or "default"), tone_map["default"])
        self.accent.setStyleSheet(f"background: {accent}; border-radius: 2px;")
        self.value_label.setStyleSheet(f"font-size: 20px; font-weight: 800; color: {value_color};")
