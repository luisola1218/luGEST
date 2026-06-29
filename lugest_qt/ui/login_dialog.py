from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
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
        self.setMinimumSize(920, 560)
        self.resize(1040, 620)
        self.setStyleSheet(
            f"""
            QDialog {{
                background: #e6edf5;
            }}
            QFrame#LoginShell {{
                background: #f8fbfe;
                border: 1px solid #b7c7d8;
                border-radius: 10px;
            }}
            QFrame#BrandPanel {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 {primary}, stop:0.62 #17365f, stop:1 #1f2937);
                border-radius: 8px;
                border: 1px solid #14315f;
            }}
            QFrame#LogoPlate {{
                background: rgba(255,255,255,0.92);
                border-radius: 8px;
                border: 1px solid rgba(255,255,255,0.34);
            }}
            QFrame#AccessPanel {{
                background: #ffffff;
                border-radius: 8px;
                border: 1px solid #c4d0df;
            }}
            QLabel#BrandName {{
                color: #f8fbff;
                font-size: 34px;
                font-weight: 950;
            }}
            QLabel#BrandSubtitle {{
                color: rgba(255,255,255,0.74);
                font-size: 13px;
                font-weight: 700;
            }}
            QLabel#CardTitle {{
                color: #0f172a;
                font-size: 28px;
                font-weight: 900;
            }}
            QLabel#CardText {{
                color: #5b6b80;
                font-size: 14px;
            }}
            QLabel#FieldLabel {{
                color: #233b56;
                font-size: 12px;
                font-weight: 800;
            }}
            QLineEdit {{
                min-height: 42px;
                padding: 0 12px;
                border: 1px solid #b8c7d9;
                border-radius: 7px;
                background: #fbfdff;
                font-size: 15px;
                color: #0f172a;
            }}
            QLineEdit:focus {{
                border: 1px solid {primary};
                background: #ffffff;
            }}
            QCheckBox {{
                color: #40546c;
                font-size: 12px;
                font-weight: 650;
            }}
            QPushButton {{
                min-height: 42px;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 800;
                padding: 0 16px;
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
                background: #f4f7fb;
                color: #173252;
                border: 1px solid #c7d6e8;
            }}
            QPushButton#SecondaryAction:hover {{
                background: #e3edf8;
            }}
            """
        )

        root = QHBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(0)

        shell = QFrame()
        shell.setObjectName("LoginShell")
        shell_layout = QHBoxLayout(shell)
        shell_layout.setContentsMargins(12, 12, 12, 12)
        shell_layout.setSpacing(12)
        root.addWidget(shell)

        logo_path = backend.logo_path

        brand_panel = QFrame()
        brand_panel.setObjectName("BrandPanel")
        brand_layout = QVBoxLayout(brand_panel)
        brand_layout.setContentsMargins(34, 34, 34, 34)
        brand_layout.setSpacing(16)

        brand = QLabel("LUGEST")
        brand.setObjectName("BrandName")
        brand_layout.addWidget(brand)

        brand_subtitle = QLabel("ERP industrial")
        brand_subtitle.setObjectName("BrandSubtitle")
        brand_layout.addWidget(brand_subtitle)
        brand_layout.addSpacing(18)

        logo_plate = QFrame()
        logo_plate.setObjectName("LogoPlate")
        logo_plate_layout = QVBoxLayout(logo_plate)
        logo_plate_layout.setContentsMargins(18, 18, 18, 18)
        logo_plate_layout.setSpacing(0)
        hero_logo = QLabel()
        hero_logo.setAlignment(Qt.AlignCenter)
        if isinstance(logo_path, Path) and logo_path.exists():
            hero_pixmap = QPixmap(str(logo_path))
            if not hero_pixmap.isNull():
                hero_logo.setPixmap(hero_pixmap.scaledToWidth(360, Qt.SmoothTransformation))
        else:
            hero_logo.setText("luGEST")
            hero_logo.setStyleSheet("font-size: 42px; font-weight: 900; color: #173252;")
        logo_plate_layout.addWidget(hero_logo, 0, Qt.AlignCenter)
        brand_layout.addWidget(logo_plate, 0)

        brand_layout.addStretch(1)

        brand_footer = QLabel("Ambiente de trabalho seguro")
        brand_footer.setStyleSheet("font-size: 12px; color: rgba(255,255,255,0.68); font-weight: 700;")
        brand_layout.addWidget(brand_footer)
        shell_layout.addWidget(brand_panel, 6)

        access_card = QFrame()
        access_card.setObjectName("AccessPanel")
        access_layout = QVBoxLayout(access_card)
        access_layout.setContentsMargins(42, 42, 42, 34)
        access_layout.setSpacing(12)
        access_layout.addStretch(1)

        title = QLabel("Entrar no luGEST")
        title.setObjectName("CardTitle")
        access_layout.addWidget(title)

        subtitle = QLabel("Usa as credenciais atribuídas ao teu posto de trabalho.")
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

        show_password = QCheckBox("Mostrar password")
        show_password.toggled.connect(
            lambda checked: self.password_edit.setEchoMode(QLineEdit.Normal if checked else QLineEdit.Password)
        )
        access_layout.addWidget(show_password)

        access_layout.addSpacing(8)

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
        shell_layout.addWidget(access_card, 5)

        self.password_edit.returnPressed.connect(self._on_login)
        self.username_edit.returnPressed.connect(self.password_edit.setFocus)
        self.username_edit.setFocus()

    def _on_login(self) -> None:
        try:
            self.backend.authenticate(self.username_edit.text(), self.password_edit.text())
        except Exception as exc:
            QMessageBox.critical(self, "Erro", str(exc))
            return
        self.accept()
