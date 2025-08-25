
import sys, os
from infra.helpers import resource_path
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QMainWindow, QGraphicsDropShadowEffect
)

from gui.Wizard import Wizard  

APP_NAME = "PlanAutomate"
SPLASH_DURATION_MS = 1200
# LOGO_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "assets", "icons", "logo.ico")
LOGO_PATH = resource_path("assets", "icons", "logo.ico")

class DotLabel(QLabel):
    def __init__(self, base_text="Démarrage...", parent=None):
        super().__init__(parent)
        self.base = base_text
        self.i = 0
        from PyQt5.QtCore import QTimer
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(280)
        self.setText(self.base)

    def _tick(self):
        self.i = (self.i + 1) % 4
        self.setText(self.base + "." * self.i)


class Splash(QWidget):
    """Fenetre de chargement de l'assistant"""
    def __init__(self):
        super().__init__(None, Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        # Card container
        card = QWidget(self)
        card.setObjectName("card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(24, 24, 24, 20)
        lay.setSpacing(10)

        logo = QLabel()
        logo.setAlignment(Qt.AlignCenter)
        if os.path.exists(LOGO_PATH):
            pm = QPixmap(LOGO_PATH)
            if not pm.isNull():
                logo.setPixmap(pm.scaledToWidth(160, Qt.SmoothTransformation))
        else:
            logo.setText(APP_NAME)
            f = QFont()
            f.setPointSize(18)
            f.setBold(True)
            logo.setFont(f)

        title = QLabel("Chargement de l'assistant...")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color:#111827;")
        title.setWordWrap(True)

        dots = DotLabel("Démarrage...")
        dots.setAlignment(Qt.AlignCenter)
        dots.setStyleSheet("color:#6B7280;")

        lay.addWidget(logo)
        lay.addSpacing(6)
        lay.addWidget(title)
        lay.addWidget(dots)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setOffset(0, 10)
        shadow.setBlurRadius(20)
        shadow.setColor(Qt.black)
        card.setGraphicsEffect(shadow)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addStretch(1)
        root.addWidget(card, alignment=Qt.AlignHCenter)
        root.addStretch(1)

        card.setFixedWidth(420)
        card.setStyleSheet("""
            #card {
                background: #FFFFFF;
                border: 1px solid #E5E7EB;
                border-radius: 18px;
            }
        """)
        self.resize(520, 280)

    def show_centered(self):
        self.show()
        scr = QApplication.primaryScreen().availableGeometry()
        self.move(
            scr.center().x() - self.width() // 2,
            scr.center().y() - self.height() // 2,
        )