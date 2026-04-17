from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)


class LoginDialog(QDialog):
    def __init__(self, backend, parent=None) -> None:
        super().__init__(parent)
        self.backend = backend
        primary = str(backend.branding.get("primary_color", "#0b1f66") or "#0b1f66")
        self.setWindowTitle("luGEST Qt")
        self.setModal(True)
        self.setMinimumSize(1080, 680)
        self.setStyleSheet(
            f"""
            QDialog {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #edf3fb, stop:1 #dce6f3);
            }}
            QFrame#HeroPanel {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 {primary}, stop:0.55 #10254d, stop:1 #091326);
                border-radius: 32px;
                border: 1px solid rgba(255,255,255,0.08);
            }}
            QFrame#LogoPlate {{
                background: rgba(255,255,255,0.10);
                border-radius: 26px;
                border: 1px solid rgba(255,255,255,0.12);
            }}
            QFrame#AccessCard {{
                background: rgba(255,255,255,0.98);
                border-radius: 28px;
                border: 1px solid #c8d6e7;
            }}
            QLabel#HeroBadge {{
                background: rgba(255,255,255,0.12);
                color: #f8fbff;
                padding: 8px 14px;
                border-radius: 999px;
                font-size: 12px;
                font-weight: 800;
            }}
            QLabel#CardEyebrow {{
                color: #37506f;
                font-size: 12px;
                font-weight: 800;
                letter-spacing: 0.5px;
            }}
            QLabel#CardTitle {{
                color: #0f172a;
                font-size: 30px;
                font-weight: 900;
            }}
            QLabel#CardText {{
                color: #52627a;
                font-size: 14px;
            }}
            QLabel#FieldLabel {{
                color: #27405d;
                font-size: 12px;
                font-weight: 800;
            }}
            QLineEdit {{
                min-height: 46px;
                padding: 0 16px;
                border: 1px solid #bbcade;
                border-radius: 16px;
                background: #f7fbff;
                font-size: 15px;
                color: #0f172a;
            }}
            QLineEdit:focus {{
                border: 2px solid {primary};
                background: #ffffff;
            }}
            QPushButton {{
                min-height: 44px;
                border-radius: 16px;
                font-size: 14px;
                font-weight: 800;
                padding: 0 18px;
            }}
            QPushButton#PrimaryAction {{
                background: {primary};
                color: #ffffff;
                border: 0;
            }}
            QPushButton#PrimaryAction:hover {{
                background: #15387d;
            }}
            QPushButton#SecondaryAction {{
                background: #edf3fa;
                color: #173252;
                border: 1px solid #c7d6e8;
            }}
            QPushButton#SecondaryAction:hover {{
                background: #e3edf8;
            }}
            """
        )

        root = QHBoxLayout(self)
        root.setContentsMargins(34, 34, 34, 34)
        root.setSpacing(28)

        logo_path = backend.logo_path

        hero = QFrame()
        hero.setObjectName("HeroPanel")
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(36, 34, 36, 34)
        hero_layout.setSpacing(18)

        hero_top = QHBoxLayout()
        brand = QLabel("luGEST")
        brand.setStyleSheet("font-size: 28px; font-weight: 900; color: #ffffff;")
        badge = QLabel("INDUSTRIAL")
        badge.setObjectName("HeroBadge")
        hero_top.addWidget(brand, 0, Qt.AlignVCenter)
        hero_top.addStretch(1)
        hero_top.addWidget(badge, 0, Qt.AlignVCenter)
        hero_layout.addLayout(hero_top)

        logo_plate = QFrame()
        logo_plate.setObjectName("LogoPlate")
        logo_plate_layout = QVBoxLayout(logo_plate)
        logo_plate_layout.setContentsMargins(24, 24, 24, 24)
        logo_plate_layout.setSpacing(0)
        hero_logo = QLabel()
        hero_logo.setAlignment(Qt.AlignCenter)
        if isinstance(logo_path, Path) and logo_path.exists():
            hero_pixmap = QPixmap(str(logo_path))
            if not hero_pixmap.isNull():
                hero_logo.setPixmap(hero_pixmap.scaledToWidth(500, Qt.SmoothTransformation))
        logo_plate_layout.addWidget(hero_logo, 0, Qt.AlignCenter)
        hero_layout.addWidget(logo_plate, 1, Qt.AlignCenter)

        hero_title = QLabel("ERP Industrial")
        hero_title.setWordWrap(True)
        hero_title.setStyleSheet("font-size: 34px; font-weight: 900; color: #ffffff; line-height: 1.1;")
        hero_title.setAlignment(Qt.AlignCenter)
        hero_layout.addWidget(hero_title)

        hero_layout.addStretch(1)

        hero_footer = QLabel("Acesso ao sistema")
        hero_footer.setStyleSheet("font-size: 13px; color: rgba(255,255,255,0.64); font-weight: 700;")
        hero_footer.setAlignment(Qt.AlignCenter)
        hero_layout.addWidget(hero_footer)
        root.addWidget(hero, 8)

        access_card = QFrame()
        access_card.setObjectName("AccessCard")
        access_layout = QVBoxLayout(access_card)
        access_layout.setContentsMargins(38, 38, 38, 38)
        access_layout.setSpacing(14)

        eyebrow = QLabel("ACESSO")
        eyebrow.setObjectName("CardEyebrow")
        eyebrow.setAlignment(Qt.AlignCenter)
        access_layout.addWidget(eyebrow)

        title = QLabel("Entrar no luGEST")
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("CardTitle")
        access_layout.addWidget(title)

        subtitle = QLabel("Introduz as credenciais para continuar no ambiente de trabalho.")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setWordWrap(True)
        subtitle.setObjectName("CardText")
        access_layout.addWidget(subtitle)

        access_layout.addSpacing(10)

        user_label = QLabel("Utilizador")
        user_label.setObjectName("FieldLabel")
        access_layout.addWidget(user_label)
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("Introduz o utilizador")
        access_layout.addWidget(self.username_edit)

        password_label = QLabel("Password")
        password_label.setObjectName("FieldLabel")
        access_layout.addWidget(password_label)
        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Introduz a password")
        self.password_edit.setEchoMode(QLineEdit.Password)
        access_layout.addWidget(self.password_edit)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        login_btn = QPushButton("Entrar")
        login_btn.setObjectName("PrimaryAction")
        login_btn.clicked.connect(self._on_login)
        cancel_btn = QPushButton("Sair")
        cancel_btn.setObjectName("SecondaryAction")
        cancel_btn.clicked.connect(self.reject)
        actions.addWidget(login_btn, 1)
        actions.addWidget(cancel_btn, 1)
        access_layout.addLayout(actions)

        access_layout.addStretch(1)
        root.addWidget(access_card, 5)

        self.password_edit.returnPressed.connect(self._on_login)
        self.username_edit.setFocus()

    def _on_login(self) -> None:
        try:
            self.backend.authenticate(self.username_edit.text(), self.password_edit.text())
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self.accept()
