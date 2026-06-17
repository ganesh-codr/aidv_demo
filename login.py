import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFrame
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

# ---------------------------------------------------------------------------
# User registry  (plain-text passwords are fine for a local demo app)
# ---------------------------------------------------------------------------
USERS = {
    "admin":    {"password": "admin",  "role": "admin"},
    "engineer": {"password": "eng",    "role": "engineer"},
    "viewer": {"password": "view",    "role": "engineer"},
}


class LoginWindow(QMainWindow):
    """
    Splash-style login window shown at application start.
    On success it closes itself and opens the main application window
    with the appropriate role restrictions applied.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("DocuChat AI — Sign In")
        self.setFixedSize(440, 560)
        self.setStyleSheet(self._stylesheet())

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setAlignment(Qt.AlignCenter)
        outer.setContentsMargins(0, 0, 0, 0)

        # ── Login card ─────────────────────────────────────────────────────
        card = QFrame()
        card.setObjectName("LoginCard")
        card.setFixedWidth(370)
        cl = QVBoxLayout(card)
        cl.setContentsMargins(40, 40, 40, 40)
        cl.setSpacing(0)

        # Brand
        brand = QLabel("AI Design Validator")
        brand.setObjectName("BrandLabel")
        brand.setAlignment(Qt.AlignCenter)
        cl.addWidget(brand)
        cl.addSpacing(6)

        subtitle = QLabel("Sign in to your workspace")
        subtitle.setObjectName("SubtitleLabel")
        subtitle.setAlignment(Qt.AlignCenter)
        cl.addWidget(subtitle)
        cl.addSpacing(32)

        # Username
        user_lbl = QLabel("Username")
        user_lbl.setObjectName("FieldLabel")
        cl.addWidget(user_lbl)
        cl.addSpacing(6)

        self.username_field = QLineEdit()
        self.username_field.setObjectName("LoginField")
        self.username_field.setPlaceholderText("Enter username")
        self.username_field.setFixedHeight(44)
        cl.addWidget(self.username_field)
        cl.addSpacing(18)

        # Password
        pass_lbl = QLabel("Password")
        pass_lbl.setObjectName("FieldLabel")
        cl.addWidget(pass_lbl)
        cl.addSpacing(6)

        self.password_field = QLineEdit()
        self.password_field.setObjectName("LoginField")
        self.password_field.setPlaceholderText("Enter password")
        self.password_field.setEchoMode(QLineEdit.Password)
        self.password_field.setFixedHeight(44)
        self.password_field.returnPressed.connect(self.attempt_login)
        cl.addWidget(self.password_field)
        cl.addSpacing(10)

        # Error label (hidden initially)
        self.error_label = QLabel("")
        self.error_label.setObjectName("ErrorLabel")
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setFixedHeight(18)
        cl.addWidget(self.error_label)
        cl.addSpacing(18)

        # Login button
        self.login_btn = QPushButton("Sign In")
        self.login_btn.setObjectName("LoginButton")
        self.login_btn.setFixedHeight(46)
        self.login_btn.clicked.connect(self.attempt_login)
        cl.addWidget(self.login_btn)
        cl.addSpacing(24)

        # Role hint strip
        hint = QLabel("admin · engineer")
        hint.setObjectName("HintLabel")
        hint.setAlignment(Qt.AlignCenter)
        cl.addWidget(hint)

        outer.addWidget(card)

    # ── Auth ────────────────────────────────────────────────────────────────

    def attempt_login(self):
        username = self.username_field.text().strip().lower()
        password = self.password_field.text()

        user = USERS.get(username)
        if user and user["password"] == password:
            from main import MainWindow
            self.main_window = MainWindow(role=user["role"], username=username)
            self.main_window.show()
            self.close()
        else:
            self.error_label.setText("Invalid username or password.")
            self.password_field.clear()
            self.password_field.setFocus()

    # ── Stylesheet ──────────────────────────────────────────────────────────

    def _stylesheet(self) -> str:
        return """
        QMainWindow, QWidget {
            background-color: #0f0f11;
        }

        QFrame#LoginCard {
            background-color: #16161a;
            border: 1px solid #2a2a35;
            border-radius: 14px;
        }

        QLabel#BrandLabel {
            font-size: 28px;
            font-weight: bold;
            color: #89b4fa;
            letter-spacing: 1px;
        }

        QLabel#SubtitleLabel {
            font-size: 13px;
            color: #6c7086;
        }

        QLabel#FieldLabel {
            font-size: 12px;
            color: #a6adc8;
            font-weight: 600;
        }

        QLineEdit#LoginField {
            background-color: #1c1c24;
            color: #eceff4;
            border: 1px solid #313244;
            border-radius: 7px;
            padding: 0px 14px;
            font-size: 13px;
        }
        QLineEdit#LoginField:focus {
            border-color: #7c4dff;
        }

        QLabel#ErrorLabel {
            color: #f38ba8;
            font-size: 11px;
            font-style: italic;
        }

        QPushButton#LoginButton {
            background-color: #7c4dff;
            color: #ffffff;
            border: none;
            border-radius: 7px;
            font-size: 14px;
            font-weight: bold;
        }
        QPushButton#LoginButton:hover  { background-color: #9d7cff; }
        QPushButton#LoginButton:pressed { background-color: #5d2bf0; }

        QLabel#HintLabel {
            font-size: 10px;
            color: #45475a;
            letter-spacing: 2px;
        }
        """


# ── Entry point (standalone testing) ────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = LoginWindow()
    win.show()
    sys.exit(app.exec())
