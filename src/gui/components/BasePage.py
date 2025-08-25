from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QLabel
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
from infra.helpers import resource_path

def create_page_layout(titre, desc):
    main_layout = QVBoxLayout()
    main_layout.setContentsMargins(0, 0, 0, 0)
    container_layout = QHBoxLayout()
    container_layout.setContentsMargins(0, 0, 0, 0)
    title_layout = QVBoxLayout()
    title_label = QLabel(titre)
    title_label.setStyleSheet("font-weight: bold;")
    desc_label = QLabel(desc)
    
    title_layout.addWidget(title_label)
    title_layout.addWidget(desc_label)

    logo = QLabel()
    icon_path = resource_path("assets", "icons", "axi.ico")
    # icon_path = os.path.join(os.path.dirname(__file__), "..","..", "assets", "icons", "axi.ico")
    icon = QIcon(icon_path)
    # logo.setPixmap(pixmap)
    logo.setPixmap(icon.pixmap(80, 80))
    logo.setContentsMargins(0, 0, 0, 0)
    # logo.align(Qt.AlignRight)
    line = QLabel()

    line.setFrameShape(QLabel.HLine)
    line.setStyleSheet("""
    QLabel {
        color:gray;
        margin-bottom: 5px;
    }""")
    container_layout.addLayout(title_layout)
    container_layout.addWidget(logo, alignment=Qt.AlignRight)
    main_layout.addLayout(container_layout)
    main_layout.addWidget(line)
    main_layout.addSpacing(3)
    # main_layout.
    
    return main_layout