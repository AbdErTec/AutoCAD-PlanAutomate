from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont 
from PyQt5.QtWidgets import QApplication

from gui.splash import Splash, APP_NAME, SPLASH_DURATION_MS
from gui.Wizard import Wizard
import sys

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion("1.0")
    app.setOrganizationName("Axigeo")

    font = QFont("Segoe UI", 11)
    app.setFont(font)

    splash = Splash()
    splash.show_centered()

    def launch_wizard():
        splash.close()
        w = Wizard()
        w.show()
        app.main_window = w

    QTimer.singleShot(SPLASH_DURATION_MS, launch_wizard)

    sys.exit(app.exec_())

