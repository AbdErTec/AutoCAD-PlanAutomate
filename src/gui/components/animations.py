# animations.py
from PyQt5.QtCore import QObject, QTimer, pyqtSlot, QPropertyAnimation
from PyQt5.QtWidgets import QLabel, QProgressBar

class StatusDots(QObject):
    def __init__(self, label: QLabel):
        super().__init__()
        self.label = label
        self.base_text = ""
        self.dots_count = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_dots)
        self.timer.start(500)

    @pyqtSlot(str)
    def set_text(self, text):
        self.base_text = text
        self.dots_count = 0  # reset
        self.update_dots()
    
    def update_dots(self):
        self.dots_count += 1
        if self.dots_count > 3:
            self.dots_count = 1
        self.label.setText(f"{self.base_text}{'.' * self.dots_count}")

class FadeText(QObject):
    def __init__(self, label: QLabel):
        super().__init__()
        self.label = label
        self.anim = QPropertyAnimation(label, b"windowOpacity")
        self.anim.setDuration(400)
        self.anim.setStartValue(0)
        self.anim.setEndValue(1)
    
    @pyqtSlot(str)
    def set_text(self, text):
        self.label.setText(text)
        self.anim.stop()
        self.anim.start()


class ProgressPulse(QObject):
    def __init__(self, progress_bar: QProgressBar):
        super().__init__()
        self.progress = progress_bar
        self.timer = QTimer()
        self.timer.timeout.connect(self.pulse)
        self.increasing = True
        self.value = 0
        self.timer.start(50)
        # self.value.setStyleSheet('margin-left: 10px;')
    
    def pulse(self):
        if self.increasing:
            self.value += 1
            if self.value >= 10:
                self.increasing = False
        else:
            self.value -= 1
            if self.value <= 0:
                self.increasing = True
        # small “glow” effect
        self.progress.setStyleSheet(f"QProgressBar::chunk {{ background-color: rgba(100, 255, 150, {150 + self.value}); padding-left: 2px; padding-right: 2px; }}")


class EasterEggLabel(QObject):
    def __init__(self, label: QLabel):
        super().__init__()
        self.label = label
        self.timer = QTimer()
        self.timer.timeout.connect(self.egg)
        self.timer.start(3000)
    
    @pyqtSlot(str)
    def set_text(self, text):
        self.label.setText(text)
    
    def egg(self):
        # tiny shake
        geom = self.label.geometry()
        self.label.move(geom.x() + 2, geom.y())
        QTimer.singleShot(100, lambda: self.label.move(geom.x(), geom.y()))
