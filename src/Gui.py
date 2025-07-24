 
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox
)
import win32com.client
import sys
import threading
import asyncio
import utils
import time
import pythoncom

class Gui(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoCAD PlanAutomate")
        self.resize(500, 300)

    
        self.file_paths = [None, None, None]

        layout = QVBoxLayout()

        for i in range(2):
            h_layout = QHBoxLayout()
            # label = QLabel(f"File {i+1}: Aucun fichier selectionne")
            label = QLabel(f"Aucun fichier selectionne")
            btn = QPushButton("Parcourir")
            btn.clicked.connect(lambda _, x=i, l=label: self.importDialogue(x, l))
            # btn.clicked.connect(lambda _, l=label: self.importDialogue(l))
            h_layout.addWidget(label)
            h_layout.addWidget(btn)
            layout.addLayout(h_layout)

        # Client info fields
        self.name_input = QLineEdit()
        self.address_input = QLineEdit()
        self.phone_input = QLineEdit()

        layout.addWidget(QLabel("Client Name:"))
        layout.addWidget(self.name_input)
        layout.addWidget(QLabel("Client Address:"))
        layout.addWidget(self.address_input)
        layout.addWidget(QLabel("Client Phone:"))
        layout.addWidget(self.phone_input)

        # Submit button
        submit_btn = QPushButton("Submit")
        submit_btn.clicked.connect(self.submit)
        layout.addWidget(submit_btn)

        self.setLayout(layout)

    def importDialogue(self, index, label):
    # app = QApplication(sys.argv)
        file_path, _ = QFileDialog.getOpenFileName(
                self,
                f"Importer fichier",
                "",
                "Fichiers DWG (*.dwg);;Tous les fichiers (*)"
            )
        # return file_path, _

        if file_path:
            self.file_paths[index] = file_path
            label.setText(f"File {index + 1} : {file_path}")

    def submit(self):
        if not self.name_input.text().strip():
            QMessageBox.warning(self, "Erreur", "Veuillez remplir toutes les informations")
            return

        def partie_b_thread_safe():
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("AutoCAD.Application")
                destination_doc = app.ActiveDocument

                print("📌 Insertion du plan en cours...")
                bbox_data = utils.inserer_plan(self.file_paths[0], destination_doc)
                time.sleep(2)

                print("📌 'Calcul du bbox'...")
                bbox_data_plan = utils.calculer_bbox(self.file_paths[0])                
                time.sleep(2)
               
                print("📌 Création du frame...")
                data_coords= utils.creer_frame_a4(bbox_data_plan)
                time.sleep(2)
                
                leg_coords = [data_coords[0], data_coords[1]]
                table_coords = [data_coords[2], data_coords[3]]
                print("📌 Insertion de la légende...")
                utils.inserer_legende(self.file_paths[1], bbox_data_plan, leg_coords)
                time.sleep(2)

                print("📌 Insertion des coordonnées...")
                utils.inserer_tableau(self.file_paths[0],bbox_data_plan, table_coords)
                time.sleep(2)

                print("✅ Tous les composants ont été insérés avec succès.")
                QMessageBox.information(self, "Success", "Tous les composants ont été insérés avec succès.")
                sys.exit(0)

            except Exception as e:
                print(f"💥 Erreur dans le thread AutoCAD: {e}")
                QMessageBox.warning(self, "Erreur", "Une Erreur est survenue lors la creation du plan")
                sys.exit(0)
            finally:
                pythoncom.CoUninitialize()

        # Thread-safe execution
        thread = threading.Thread(target=partie_b_thread_safe)
        thread.start()
        QMessageBox.information(self, "Traitement lancé", "Insertion en cours dans AutoCAD... Vous serez notifié.")
        thread.join()
        
def start():
    app = QApplication(sys.argv)
    window = Gui()
    window.show()
    sys.exit(app.exec())

