from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

class Field(QWidget):
    def __init__(self, label_1, input1_type=QLineEdit, label_2=None, input2_type=QLineEdit, isar=True):
        super().__init__()
        layout = QFormLayout()

        # Create a container for inputs
        input_container = QWidget()
        input_layout = QHBoxLayout(input_container)
        input_layout.setSpacing(10)

        # Input 1
        self.input_1 = input1_type()
        input_layout.addWidget(self.input_1)

        if label_2:
            self.input_1.setFixedWidth(200)

            input_layout.addSpacing(20)
            input_layout.addStretch()

            # Input 2
            
            if isar:
                # Arabic label to the right of input
                self.input_2 = input2_type()
                if input2_type == QLineEdit:
                    input2_type.setLayoutDirection(self.input_2, Qt.RightToLeft)
                    input2_type.setAlignment(self.input_2, Qt.AlignRight)
                self.input_2.setFixedWidth(200)
                input_layout.addWidget(self.input_2)

                input_layout.addWidget(QLabel(f"{label_2}:"))
            else:
                # English label to the left of input
                input_layout.addWidget(QLabel(f"{label_2}:"))
                self.input_2 = input2_type()
                self.input_2.setFixedWidth(200)
                input_layout.addWidget(self.input_2)

        # Add to form layout
        layout.addRow(f"{label_1}:", input_container)
        layout.setContentsMargins(0, 0, 0, 0)

        self.setLayout(layout)

    def get_value(self):
        return self.input_1.text()

    def get_second_value(self):
        if hasattr(self, 'input_2'):
            return self.input_2.text()
        return None

    @staticmethod
    def create_groupbox(form_layout, title = None):
        form_group = QGroupBox(title)
        form_group.setLayout(form_layout)
        form_group.setStyleSheet("""
            QGroupBox {
                background-color: #f7f7f7;
                font-weight: bold;
                border: 1px solid gray;
                border-radius: 8px;
                margin-top: 10px;
            }`
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
            }
        """)

        return form_group