try:
    import ctypes
    from ctypes import wintypes
    myappid = 'axigeo.planautomate.1.0' 
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass  

from PyQt5.QtWidgets import *
from PyQt5.QtGui import QFont,QPixmap, QIcon
from PyQt5.QtCore import Qt
import os
import re
from pathlib import Path

from workers.autocad_prepworker import AutoCADPrepWorker
from workers.autocad_worker import AutoCADWorker

from infra.com_manager import COMManager as CM
from infra.helpers import resource_path

from gui.components.BasePage import create_page_layout as baseLayout
from gui.components.Field import Field
from gui.components.animations import StatusDots, FadeText, ProgressPulse, EasterEggLabel

class Wizard(QMainWindow):
    def __init__(self):
        super().__init__()
        #initialize window
        self.setWindowTitle("Plan Automate")
        # icon_path = os.path.join(os.path.dirname(__file__), "logo.png")
        # icon_path = os.path.join(os.path.dirname(__file__), "..",'..', "assets",'icons', "logo.ico")
        icon_path = resource_path("assets", "icons", "logo.ico")

        self.setWindowFlags(Qt.Window | Qt.CustomizeWindowHint | Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        icon = QIcon(icon_path)
        self.setWindowIcon(icon)
        QApplication.instance().setWindowIcon(icon)
        self.resize(1000, 600)

        self.setWindowIcon(QIcon(icon_path))

        central_widget = QWidget()
        central_widget.setContentsMargins(0, 0, 0, 0)
        self.setCentralWidget(central_widget)
        
        self.client_data = {}
        self.images_paths = [None, None]
        self.file_paths = [None, None, None]
        self.worker = None

        self.stack = QStackedLayout()
        self.onboarding_page()
        self.init_page1()
        self.init_page2()
        self.init_page3()
        self.wait_page()
        self.prep_page()
        self.init_page4()
        self.offboarding_page()

        self.stack.setContentsMargins(0, 0, 0, 0)
        
        container = QWidget()
        container.setLayout(self.stack)

        layout = QVBoxLayout()
        layout.addWidget(container)
        central_widget.setLayout(layout)

    def onboarding_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 15, 0) 
        #===right side===
        title_layout = QVBoxLayout()
        title = QLabel("Bienvenue dans AutoCAD Plan Automate")
        title.setFont(QFont("Segoe UI", 18))
        title.setStyleSheet("margin-top:50px; font-weight: Bold")
        title_layout.addWidget(title)
        desc = QLabel("Cet assistant vous guidera dans la saisie les données de la cartouche, l'importation de fichiers et la génération de vos plans AutoCAD.\n\n"
                           "Cliquez \"Suivant\" pour continuer, ou \"Annuler\" pour quitter l'assistant.")
        desc.setWordWrap(True)
        title_layout.addWidget(desc)
        
        title_layout.setAlignment(Qt.AlignLeft)
        title_layout.setContentsMargins(5, 0, 0, 5)

        #===Left side===
        logo_layout = QVBoxLayout()
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo = QLabel()
        icon_path = resource_path("assets", "icons", "axi.png")
        pixmap = QPixmap(icon_path) 
        # pixmap = QPixmap(os.path.join(os.path.dirname(__file__), "..","..", "assets",'icons', "axi.png"))
        logo.setPixmap(pixmap.scaled(350, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        logo_layout.addWidget(logo)
        logo_layout.itemAt(0).widget().setStyleSheet("background-color: white;")
        logo_layout.setAlignment(Qt.AlignLeft)
        content_layout.addLayout(logo_layout)

        content_layout.addLayout(title_layout)
        content_layout.setStretch(0,3)
        content_layout.setStretch(1,7)

        layout.addLayout(content_layout)

        # === button layout ===
        btn_layout = QHBoxLayout()
        next_btn = QPushButton("Suivant>>")
        next_btn.setFixedWidth(100)
        next_btn.clicked.connect(self.goto_page1)
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setFixedWidth(100)
        cancel_btn.clicked.connect(self.close)
        btn_layout.addWidget(next_btn)
        btn_layout.addSpacing(10)
        btn_layout.addWidget(cancel_btn)
        btn_layout.setAlignment(Qt.AlignRight)

        #===credit===
        credit = QLabel("Abderrahmane Techa - 2025")
        credit.setAlignment(Qt.AlignLeft)
        credit.setFixedHeight(20)
        credit.setContentsMargins(0, 5, 0, 0)
        credit.setStyleSheet("color: gray; font-size: 10px; margin-top: 5px;")

        #===add Layouts===
        layout.addWidget(credit) 
        layout.addLayout(btn_layout)

        # ===Final layout===
        page.setLayout(layout)
        self.stack.addWidget(page)
    
    def init_page1(self):
        """Page Cartouche"""
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        title_layout = baseLayout("Cartouche", "Veuillez  remplir les informations à intégrer dans le cartouche du plan")
        layout.addLayout(title_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # === group 1 ===
        form_layout1 = QFormLayout()
        form_layout1.setSpacing(5)

        self.region_field = Field(label_1="Région *", label_2="الجهة *")
        form_layout1.addRow(self.region_field)

        self.province_field = Field(label_1="Province/Préfecture *", label_2= "الإقليم/العمالة *")
        form_layout1.addRow(self.province_field)

        self.commune_field = Field(label_1="Commune *", label_2="الجماعة *")
        form_layout1.addRow(self.commune_field)

        self.situation_field = Field(label_1="Situation *", label_2="الموقع *")
        form_layout1.addRow(self.situation_field)

        form_group1 = Field.create_groupbox(form_layout1)
        scroll_layout.addWidget(form_group1)
        
        # === group 2 ===
        form_layout2 = QFormLayout()
        form_layout2.setSpacing(5)

        self.plan_field = Field("Plan *")
        form_layout2.addRow(self.plan_field)

        form_group2 = Field.create_groupbox(form_layout2)
        scroll_layout.addWidget(form_group2)

        # === group 3 ===
        form_layout3 = QFormLayout()
        form_layout3.setSpacing(5)

        self.contenance_field = Field("Contenance *")
        form_layout3.addRow(self.contenance_field)

        form_group3 = Field.create_groupbox(form_layout3)
        scroll_layout.addWidget(form_group3)

        # === group 4 ===
        form_layout4 = QFormLayout()
        form_layout4.setSpacing(5)

        self.demande_par_field = Field("Plan demandé par *")
        form_layout4.addRow(self.demande_par_field)

        self.propriete_field = Field("Proprieté dte")
        form_layout4.addRow(self.propriete_field)

        self.ref_field = Field("Référence Fonciere")
        form_layout4.addRow(self.ref_field)

        self.observations_field = Field("Observations")
        form_layout4.addRow(self.observations_field)

        form_group4 = Field.create_groupbox(form_layout4)
        scroll_layout.addWidget(form_group4)
        
        # === group 5 ===
        form_layout5 = QFormLayout()
        form_layout5.setSpacing(5)

        self.declaration_par_field = Field(label_1="Déclaration de *", label_2="CIN *", isar=False)
        form_layout5.addRow(self.declaration_par_field)

        form_group5 = Field.create_groupbox(form_layout5)
        scroll_layout.addWidget(form_group5)
        

        # === group 6 ===
        form_layout6 = QFormLayout()
        form_layout6.setSpacing(5)

        self.levage_field = Field("Levé le *", input1_type=QDateEdit, label_2= "Agent leveur *", isar=False)
        form_layout6.addRow(self.levage_field)

        self.echelle_field = Field("Echelle *")
        form_layout6.addRow(self.echelle_field)

        self.date_num_dossier_field = Field(label_1="Date *",input1_type=QDateEdit,label_2="N° Dossier *",isar=False)
        form_layout6.addRow(self.date_num_dossier_field)

        self.nivellement_field = Field("Nivellement")
        form_layout6.addRow(self.nivellement_field)

        self.coords_field = Field("Coordonnées *")
        form_layout6.addRow(self.coords_field)

        self.fichier_xref_field = Field(label_1="Fichier", label_2="xref", isar=False)
        form_layout6.addRow(self.fichier_xref_field)

        form_group6 = Field.create_groupbox(form_layout6)
        scroll_layout.addWidget(form_group6)

        # === group 7 ===
        form_layout7 = QFormLayout()
        form_layout7.setSpacing(5)

        self.img_labels = []

        self.apercu_fond_field = Field(label_1="Aperçu sur fond haut *", label_2="Aperçu sur fond bas *", isar=False)
        form_layout7.addRow(self.apercu_fond_field)

        logo_qr_widget = QWidget()
        logo_qr_layout = QHBoxLayout(logo_qr_widget)

        # --- Logo ---
        logo_layout = QHBoxLayout()
        logo_label = QLabel("Logo:  Aucun fichier sélectionné")
        self.img_labels.append(logo_label)
        btn_logo = QPushButton("Parcourir..")
        btn_logo.setFixedWidth(100)
        btn_logo.clicked.connect(lambda: self.import_file(0,"img"))
        logo_layout.addWidget(logo_label)
        logo_layout.addWidget(btn_logo)
        logo_layout.addStretch()
        
        # --- QR ---
        qr_layout = QHBoxLayout()
        qr_label = QLabel("QR Code:  Aucun fichier sélectionné")
        self.img_labels.append(qr_label)
        btn_qr = QPushButton("Parcourir..")
        btn_qr.setFixedWidth(100)
        btn_qr.clicked.connect(lambda: self.import_file(1,"img"))
        qr_layout.addWidget(qr_label)
        qr_layout.addWidget(btn_qr)

        logo_qr_layout.addLayout(logo_layout)
        logo_qr_layout.addLayout(qr_layout)

        form_layout7.addRow(logo_qr_widget)

        form_group7 = Field.create_groupbox(form_layout7)
        scroll_layout.addWidget(form_group7)

        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        # === button layout ===
        btn_layout = QHBoxLayout()
        next_btn = QPushButton("Suivant>>")
        next_btn.setFixedWidth(100)
        next_btn.clicked.connect(self.go_to_page2)
        back_btn = QPushButton("<<Précedent")
        back_btn.setFixedWidth(100)
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setFixedWidth(100)
        cancel_btn.clicked.connect(self.close)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(next_btn)
        btn_layout.addSpacing(10)
        btn_layout.addWidget(cancel_btn)
        btn_layout.setAlignment(Qt.AlignRight)
        layout.addLayout(btn_layout)
        
        # ===Final layout===
        page.setLayout(layout)
        self.stack.addWidget(page)

    def init_page2(self):
        """selection des fichiers"""
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        self.labels = []

        title_layout = baseLayout("Sélection des fichiers", "Veuillez sélectionner les fichiers DWG qui composent votre plan (Plan, Légende, et ortophoto + carte)")
        layout.addLayout(title_layout)
        title_layout.addStretch()
        content_layout = QVBoxLayout()

        line = [None, None, None]

        for i in range(len(line)):
            l = QLabel()
            l.setFrameShape(QLabel.HLine)
            l.setStyleSheet("""
                QLabel {
                    color: gray;
                    margin-bottom: 3px;
                }
            """)
            line[i] = l # now you're saving it properly
        
        # File 1 selection
        h_layout1 = QHBoxLayout()
        label1 = QLabel("Plan principal: Aucun fichier sélectionné")
        self.labels.append(label1)
        btn1 = QPushButton("Parcourir..")
        btn1.setFixedWidth(100)
        btn1.clicked.connect(lambda: self.import_file(0))
        h_layout1.addWidget(label1)
        h_layout1.addWidget(btn1)
        h_layout1.setSpacing(10)
        content_layout.addLayout(h_layout1)

        content_layout.addWidget(line[0])
        

        h_layout2 = QHBoxLayout()
        label2 = QLabel("Légende: Aucun fichier sélectionné")
        self.labels.append(label2)
        btn2 = QPushButton("Parcourir...")
        btn2.setFixedWidth(100)
        btn2.clicked.connect(lambda: self.import_file(1))
        h_layout2.addWidget(label2)
        h_layout2.addWidget(btn2)
        h_layout2.setSpacing(10)
        content_layout.addLayout(h_layout2)

        content_layout.addWidget(line[1])

        h_layout3 = QHBoxLayout()
        label3 = QLabel("Input: Aucun fichier sélectionné")
        self.labels.append(label3)
        btn3 = QPushButton("Parcourir...")
        btn3.setFixedWidth(100)
        btn3.clicked.connect(lambda: self.import_file(2))
        h_layout3.addWidget(label3)
        h_layout3.addWidget(btn3)
        h_layout3.setSpacing(10)
        content_layout.addLayout(h_layout3)

        content_layout.addWidget(line[2])

        content_layout.setContentsMargins(10, 10, 10, 5)
        layout.addLayout(content_layout)
       
       # === button layout ===
        btn_layout = QHBoxLayout()
        next_btn = QPushButton("Suivant>>")
        next_btn.setFixedWidth(100)
        next_btn.clicked.connect(self.go_to_page3)
        back_btn = QPushButton("<<Précedent")
        back_btn.setFixedWidth(100)
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setFixedWidth(100)
        cancel_btn.clicked.connect(self.close)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(next_btn)
        btn_layout.addSpacing(10)
        btn_layout.addWidget(cancel_btn)
        btn_layout.setAlignment(Qt.AlignRight)
        layout.addLayout(btn_layout)
        
        # ===Final layout===
        page.setLayout(layout)
        self.stack.addWidget(page)

    def init_page3(self):
        """Review page"""
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)


        title_layout = baseLayout("Récapitulatif", "Vérifiez les informations ci-dessus avant de continuer.")
        layout.addLayout(title_layout)
        # # --- Scrollable area ---

        self.review_imports = QLabel()
        layout.addWidget(self.review_imports)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # # === form layout ===
        self.review_cartouche = QLabel()
        scroll_layout.addWidget(self.review_cartouche)

        scroll_area.setWidget(scroll_content)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)
        scroll_layout.addStretch()

       # === button layout ===
        btn_layout = QHBoxLayout()
        next_btn = QPushButton("Suivant>>")
        next_btn.setFixedWidth(100)
        next_btn.clicked.connect(self.start_prep_process)
        back_btn = QPushButton("<<Précedent")
        back_btn.setFixedWidth(100)
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(2))
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setFixedWidth(100)
        cancel_btn.clicked.connect(self.close)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(next_btn)
        btn_layout.addSpacing(10)
        btn_layout.addWidget(cancel_btn)
        btn_layout.setAlignment(Qt.AlignRight)
        layout.addLayout(btn_layout)
        
        # ===Final layout===
        page.setLayout(layout)
        self.stack.addWidget(page)

    def wait_page(self):
        """page de preparation (recuperation des calques et calcul des echelles)"""
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        title_layout = baseLayout("Préparation AutoCAD...", "Veuillez patienter pendant la préparation.")
        layout.addLayout(title_layout)

        content_layout = QVBoxLayout()

        self.prep_progress = QProgressBar()
        self.prep_progress.setRange(0, 0)  # Indeterminate progress (spinning)
        content_layout.addWidget(self.prep_progress)

        self.prep_status_label = QLabel("Initialisation")
        content_layout.addWidget(self.prep_status_label)
        self.prep_dots = StatusDots(self.prep_status_label)

        content_layout.setContentsMargins(10, 10, 10, 5)
        content_layout.addStretch()

        layout.addLayout(content_layout)

        self.prep_result_layout = QHBoxLayout()

        self.prep_next_btn = QPushButton("Suivant>>")
        self.prep_next_btn.setFixedWidth(100)
        self.prep_next_btn.clicked.connect(self.go_to_page4)
        self.prep_next_btn.setVisible(False)

        self.prep_retry_btn = QPushButton("Réessayer")
        self.prep_retry_btn.setFixedWidth(100)
        self.prep_retry_btn.clicked.connect(lambda: self.stack.setCurrentIndex(3))  # or relevant page
        self.prep_retry_btn.setVisible(False)

        self.prep_result_layout.addWidget(self.prep_next_btn)
        self.prep_result_layout.addWidget(self.prep_retry_btn)
        self.prep_result_layout.setAlignment(Qt.AlignRight)

        layout.addLayout(self.prep_result_layout)

        page.setLayout(layout)
        self.stack.addWidget(page)

    def prep_page(self):
        """selection des calques a inclure et choix de l'echelle"""
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        title_layout = baseLayout("Préférences du plan", "Choisissez les calques à inclure et l’échelle du plan.")
        layout.addLayout(title_layout)

        self.layer_checkboxes = {}
        self.layer_group = QVBoxLayout()  
        group_box = QGroupBox("Calques à inclure")
        group_box.setLayout(self.layer_group)  

        available_layers = []
        for layer_name in available_layers:
            cb = QCheckBox(layer_name)
            cb.setChecked(True)
            self.layer_checkboxes[layer_name] = cb
            self.layer_group.addWidget(cb)

        layout.addWidget(group_box)

        echelle_layout = QHBoxLayout()
        echelle_label = QLabel("Échelle:")
        self.echelle_dropdown = QComboBox()
        available_echelles = []
        self.echelle_dropdown.addItems(available_echelles)
        echelle_layout.addWidget(echelle_label)
        echelle_layout.addWidget(self.echelle_dropdown)
        echelle_layout.addStretch()

        layout.addLayout(echelle_layout)
        layout.addStretch()


        # === button layout ===
        btn_layout = QHBoxLayout()
        next_btn = QPushButton("Lancer l'insertion")

        next_btn.clicked.connect(self.go_to_page4)
        back_btn = QPushButton("<<Précedent")
        back_btn.setFixedWidth(100)

        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(3))
        cancel_btn = QPushButton("Annuler")
        cancel_btn.setFixedWidth(100)
        cancel_btn.clicked.connect(self.close)
        btn_layout.addWidget(back_btn)
        btn_layout.addWidget(next_btn)
        btn_layout.addSpacing(10)
        btn_layout.addWidget(cancel_btn)
        btn_layout.setAlignment(Qt.AlignRight)

        layout.addLayout(btn_layout)

        page.setLayout(layout)
        self.stack.addWidget(page)

    def init_page4(self):
        """Progress page"""
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)


        title_layout = baseLayout("Insertion en cours...", "Ne touchez pas AutoCAD lors de la generation du plan.")
        layout.addLayout(title_layout)
        content_layout = QVBoxLayout()
        
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.pulse = ProgressPulse(self.progress)
        
        content_layout.addWidget(self.progress)
        self.progress.valueChanged.connect(self.check_progress)
        
        self.status_label = QLabel("Préparation")
        content_layout.addWidget(self.status_label)
        self.dots = StatusDots(self.status_label)
        self.fade = FadeText(self.status_label)
        self.egg = EasterEggLabel(self.status_label)
        
        content_layout.setContentsMargins(10, 10, 10, 5)
        content_layout.addStretch()

        layout.addLayout(content_layout)
        # Buttons (initially hidden)
        self.result_layout = QHBoxLayout()
        self.finish_btn = QPushButton("Terminer")
        self.finish_btn.setFixedWidth(100)
        self.finish_btn.clicked.connect(self.close)
        self.finish_btn.setVisible(False)
        
        self.retry_btn = QPushButton("Réessayer")
        self.retry_btn.setFixedWidth(100)
        self.retry_btn.clicked.connect(lambda: self.stack.setCurrentIndex(3))
        self.retry_btn.setVisible(False)
        
        self.result_layout.addWidget(self.retry_btn)
        self.result_layout.addWidget(self.finish_btn)
        self.result_layout.setAlignment(Qt.AlignRight)
        layout.addLayout(self.result_layout)

        page.setLayout(layout)
        self.stack.addWidget(page)

    def offboarding_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)  
        
        #===right side===
        title_layout = QVBoxLayout()
        title = QLabel("AutoCAD Plan Automate terminé.")
        title.setFont(QFont("Segoe UI", 20))

        title_layout.addWidget(title)
        desc = QLabel("Votre plan a bien été crée. Vous pouvez maintenant fermer ce programme.\n")
        desc.setWordWrap(True)
        
        title_layout.addWidget(desc)
        title_layout.setAlignment(Qt.AlignLeft)
        title_layout.setContentsMargins(5, 0, 0, 5)

        #===Left side===
        logo_layout = QVBoxLayout()
        logo_layout.setContentsMargins(0, 0, 0, 0)

        logo = QLabel()
        pixmap = QPixmap(os.path.join(os.path.dirname(__file__), "..","..", "assets","icons", "axi.png"))
        logo.setPixmap(pixmap.scaled(350, 300, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        logo_layout.addWidget(logo)
        logo_layout.itemAt(0).widget().setStyleSheet("background-color: white;")
        logo_layout.setAlignment(Qt.AlignLeft)

        content_layout.addLayout(logo_layout)
        content_layout.addLayout(title_layout)
        content_layout.setStretch(0,3)
        content_layout.setStretch(1,7)

        layout.addLayout(content_layout)

        # === button layout ===
        self.finish_btn = QPushButton("Terminer")
        self.finish_btn.setFixedWidth(100)
        self.finish_btn.clicked.connect(self.close)
        btn_layout = QHBoxLayout()
        btn_layout.setAlignment(Qt.AlignRight)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        #===credit===
        credit = QLabel("Abderrahmane Techa - 2025")
        credit.setAlignment(Qt.AlignLeft)
        credit.setFixedHeight(20)
        # credit.setContentsMargins(0, 0, 640, 0)
        credit.setStyleSheet("color: gray; font-size: 10px; margin-top: 5px;")
        
        #===add Layouts===
        btn_layout.addWidget(credit) 
        btn_layout.addStretch()
        btn_layout.addWidget(self.finish_btn)        
        layout.addLayout(btn_layout)

        # ===Final layout===
        page.setLayout(layout)
        self.stack.addWidget(page)    
    
    def goto_page1(self):
        backup_dir = Path.home()/"Documents"/"PlanAutomate"/"backup"
        hasFiles = False
        if backup_dir.exists() and any (backup_dir.iterdir()):
            for f in backup_dir.iterdir(): 
                if "checkpoint" in str(f).lower():
                    hasFiles = True
                    break

            reply=None
            if hasFiles:
            # pass
                reply = QMessageBox.question(
                self, 
                "Plan inachevé détecté", 
                "Vous avez un plan inachevé. Voulez-vous le reprendre?\n\n"
                "Oui - Reprendre le plan existant\n"
                "Non - Commencer un nouveau plan",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                self.resume_from_backup()
                return
            else:
                # Clear backup folder
                self.clear_backup_folder(backup_dir)
                
        self.stack.setCurrentIndex(1)
    
    def convert_contenance(self, cont):
        """convertir contenance du m2 au ha, a, ca"""
        cont = int(cont)  
        hectares = cont // 10000
        remainder = cont % 10000
        
        ares = remainder // 100
        centiares = remainder % 100
        return f"{hectares} ha {ares} a {centiares} ca"

    def go_to_page2(self):
        """cartouche regex"""

        self.client_data = {
            "region": self.region_field.get_value(),
            "region_ar": self.region_field.get_second_value(),
            
            "province": self.province_field.get_value(),
            "province_ar": self.province_field.get_second_value(),

            "commune": self.commune_field.get_value(),
            "commune_ar": self.commune_field.get_second_value(),

            "situation": self.situation_field.get_value(),
            "situation_ar": self.situation_field.get_second_value(),

            "plan": self.plan_field.get_value(),

            "contenance": self.convert_contenance(self.contenance_field.get_value()) if self.contenance_field.get_value().isdigit() else None,

            "demande_par": self.demande_par_field.get_value(),
            "propriete_dte": self.propriete_field.get_value() if self.propriete_field.get_value() else '*********',
            "reference_fonciere": self.ref_field.get_value() if self.ref_field.get_value() else '*********',
            "observations": self.observations_field.get_value() if self.observations_field.get_value() else '*********',

            "declaration_par": self.declaration_par_field.get_value() if self.declaration_par_field.get_value().isalpha else None,
            "cin": self.declaration_par_field.get_second_value().upper() if len(self.declaration_par_field.get_second_value()) >= 7 and self.declaration_par_field.get_second_value().isalnum() else None,

            "leve_le": self.levage_field.get_value(),
            "agent_leveur": self.levage_field.get_second_value(),

            "echelle": self.echelle_field.get_value() if re.sub(r'[^0-9\-/]', '', self.echelle_field.get_value()) else None,

            "date": self.date_num_dossier_field.get_value(),
            "numero_dossier": self.date_num_dossier_field.get_second_value() if re.sub(r'[^0-9\-/]', '', self.date_num_dossier_field.get_second_value()) else None,
            "nivellement": self.nivellement_field.get_value() if self.nivellement_field.get_value() else '*********',
            "coordonnees": self.coords_field.get_value(),

            "fichier": self.fichier_xref_field.get_value() if self.fichier_xref_field.get_value() else '*********',
            "xref": self.fichier_xref_field.get_second_value() if self.fichier_xref_field.get_second_value() else '*********',

            "apercu_fond_haut": self.apercu_fond_field.get_value(),
            "apercu_fond_bas": self.apercu_fond_field.get_second_value(),

            "logo": self.images_paths[0],
            "qr": self.images_paths[1],
        }
        if not hasattr(self, 'client_data') or not self.client_data:
            missing_keys = []
        else:
            missing_keys = [key for key, value in self.client_data.items() if not value]

        if not hasattr(self, 'client_data') or not self.client_data or missing_keys:
            QMessageBox.warning(self, "Erreur", f"Veuillez remplir toutes les valeurs requises.\n\n (Champs manquants: {', '.join(missing_keys).replace('_', ' ')})")
            print(missing_keys)
            return

        self.stack.setCurrentIndex(2)

    def go_to_page3(self):
        """Validation des fichiers"""
        import filecmp
        
        if not self.file_paths[0] or not self.file_paths[1] or not self.file_paths[2]:
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner les trois fichiers requis.")
            return
        if filecmp.cmp(self.file_paths[0], self.file_paths[1]) or filecmp.cmp(self.file_paths[0], self.file_paths[2]) or filecmp.cmp(self.file_paths[1], self.file_paths[2]):
            QMessageBox.warning(self, "Attention", "Veuillez sélectionner trois fichiers distincts.")
        # Check if files exist
        for i, path in enumerate(self.file_paths):
            if not os.path.exists(path):
                QMessageBox.warning(self, "Erreur", f"Le fichier {i+1} n'existe plus: {path}")
                return

        self.update_review_content()
        self.stack.setCurrentIndex(3)

    def go_to_page4(self):
        self.handle_layer_selection_next()

    def check_progress(self):
        if self.progress.value() == 100:
            self.stack.setCurrentIndex(7)

    def handle_layer_selection_next(self):
        selected_layers = [
            name for name, cb in self.layer_checkboxes.items() if cb.isChecked()
        ]
        selected_echelle = self.echelle_dropdown.currentText()

        print("Selected layers:", selected_layers)
        print("Selected échelle:", selected_echelle)
    
        self.stack.setCurrentIndex(6)
        self.start_autocad_process(selected_layers, selected_echelle, self.client_data, self.images_paths)

    def update_review_content(self):
        """mettre a jour review page"""
        imports = f"""
<b>Fichiers sélectionnés:</b><br>
Plan principal: {os.path.basename(self.file_paths[0]) if self.file_paths[0] else 'N/A'}<br>
Légende: {os.path.basename(self.file_paths[1]) if self.file_paths[1] else 'N/A'}<br>
Input: {os.path.basename(self.file_paths[2]) if self.file_paths[2] else 'N/A'}<br>
        """
        self.review_imports.setText(imports)
        cartouche = f"""
    <table width='100%' cellspacing='10'>
        <tr>
            <td align='left'><b>Région:</b> {self.client_data['region']}</td>
            <td align='right'><b>الجهة:</b> {self.client_data['region_ar']}</td>
        </tr>
        <tr>
            <td align='left'><b>Province:</b> {self.client_data['province']}</td>
            <td align='right'><b>الإقليم:</b> {self.client_data['province_ar']}</td>
        </tr>
        <tr>
            <td align='left'><b>Commune:</b> {self.client_data['commune']}</td>
            <td align='right'><b>الجماعة:</b> {self.client_data['commune_ar']}</td>
        </tr>
        <tr>
            <td align='left'><b>Situation:</b> {self.client_data['situation']}</td>
            <td align='right'><b>الموقع:</b> {self.client_data['situation_ar']}</td>
        </tr>

        <tr><td colspan="2"><hr></td></tr>
        
        <tr><td align='left'><b>Plan:</b> {self.client_data['plan']}</td></tr>
        <tr><td align='left'><b>Contenance:</b> {self.client_data['contenance']}</td></tr>
        <tr><td align='left'><b>Plan demandé par:</b> {self.client_data['demande_par']}</td></tr>
        <tr><td align='left'><b>Propriété dte:</b> {self.client_data['propriete_dte']}</td></tr>
        <tr><td align='left'><b>Réference foncière:</b> {self.client_data['reference_fonciere']}</td></tr>
        <tr><td align='left'><b>Observations:</b> {self.client_data['observations']}</td></tr>

        <tr><td colspan="2"><hr></td></tr>

        <tr>
            <td align='left'><b>Déclaration de :</b> {self.client_data['declaration_par']}</td>
            <td align='left'><b>CIN:</b> {self.client_data['cin']}</td>
        </tr>

        <tr><td colspan="2"><hr></td></tr>
        
        <tr>
            <td align='left'><b>Levé le :</b> {self.client_data['leve_le']}</td>
            <td align='left'><b>Agent leveur :</b> {self.client_data['agent_leveur']}</td>
        </tr>
        <tr><td align='left'><b>Echelle:</b> {self.client_data['echelle']}</td></tr>
        <tr><td align='left'><b>Plan demandé par:</b> {self.client_data['demande_par']}</td></tr>
        <tr>
            <td align='left'><b>Date:</b> {self.client_data['date']}</td>
            <td align='left'><b>N° Dossier:</b> {self.client_data['numero_dossier']}</td>
        </tr>
        <tr><td align='left'><b>Nivellement:</b> {self.client_data['nivellement']}</td></tr>
        <tr><td align='left'><b>Coordonnées:</b> {self.client_data['coordonnees']}</td></tr>
        <tr>
            <td align='left'><b>Fichier:</b> {self.client_data['fichier']}</td>
            <td align='left'><b>Xref:</b> {self.client_data['xref']}</td>
        </tr>

        <tr><td colspan="2"><hr></td></tr>

        <tr>
            <td align='left'><b>Aperçu sur fond haut :</b> {self.client_data['apercu_fond_haut']}</td>
            <td align='left'><b>Aperçu sur fond bas :</b> {self.client_data['apercu_fond_bas']}</td>
        </tr>
        <tr>
            <td align='left'><b>Logo :</b> {self.client_data['logo']}</td>
            <td align='left'><b>Code QR :</b> {self.client_data['qr']}</td>
        </tr>
    </table>
    """
        self.review_cartouche.setText(cartouche)
        
    def import_file(self, index, type='dwg'):
        """Import des fichies DWG et images"""
        if type == "img":
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Importer un fichier", "", 
                "Images (*.jpeg, *.jpg, *.png);;Tous les fichiers (*)")
            if file_path:
                self.images_paths[index] = file_path
                filename = os.path.basename(file_path)
                prefix = "Logo: " if index == 0 else "Code QR: "
                self.img_labels[index].setText(f"{prefix}{filename}")
        
        else:    
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Importer un fichier", "", 
                "Fichiers DWG (*.dwg);;Tous les fichiers (*)")
            if file_path:
                self.file_paths[index] = file_path
                filename = os.path.basename(file_path)
                # prefix = "Plan principal: " if index == 0 else (if index == 1 "Légende: " else "Input"
                match index:
                    case 0:
                        prefix = "Plan principal: "
                    case 1:
                        prefix = "Legende: "
                    case 2:
                        prefix = "Input: "
                self.labels[index].setText(f"{prefix}{filename}")
        
    def start_prep_process(self):
            """commencer prep worker"""
            self.stack.setCurrentIndex(4)
            self.prep_progress.setValue(0)
            self.status_label.setText("Initialisation...")

            self.prep_next_btn.setVisible(False)
            self.prep_retry_btn.setVisible(False)

            self.worker = AutoCADPrepWorker(self.file_paths, self.client_data['echelle'])
            self.worker.done.connect(self.on_prep_done, type=Qt.QueuedConnection)
            self.worker.status_updated.connect(self.prep_dots.set_text, type=Qt.QueuedConnection)

            self.worker.error.connect(self.on_error, type=Qt.QueuedConnection)
            self.worker.start()

    def on_prep_done(self, data, message):
           
            print('we\'re here')
            self.layers = data.get('layers', [])
            self.echelles = data.get('echelles_finales', [])
            
            for cb in self.layer_checkboxes.values():
                cb.setParent(None)
            self.layer_checkboxes.clear()
            
            for layer in self.layers:
                if isinstance(layer, dict):
                    layer_name = layer.get('name', str(layer))  
                else:
                    layer_name = str(layer)  
                cb = QCheckBox(layer_name)
                cb.setChecked(True)
                self.layer_checkboxes[layer_name] = cb
                self.layer_group.addWidget(cb)  
            
            self.echelle_dropdown.clear()
            for echelle in self.echelles:
                if isinstance(echelle, dict):
                    echelle_name = echelle.get('label', str(echelle)) 
                else:
                    echelle_name = str(echelle)  
                self.echelle_dropdown.addItem(echelle_name)
            self.stack.setCurrentIndex(5)
            self.prep_next_btn.setVisible(True)

    def start_autocad_process(self, selected_layers, selected_echelle, client_data, images_paths):
        """comencer worker principal"""
        self.progress.setValue(0)
        self.status_label.setText("Initialisation...")
        self.finish_btn.setVisible(False)
        self.retry_btn.setVisible(False)
        
        self.worker = AutoCADWorker(self.file_paths, selected_layers, selected_echelle, client_data, self.images_paths)
        self.worker.progress_updated.connect(self.progress.setValue)
        self.worker.status_updated.connect(self.status_label.setText)
        self.worker.process_finished.connect(self.on_process_finished)

        self.worker.status_updated.connect(self.dots.set_text)
        self.worker.status_updated.connect(self.fade.set_text)
        self.worker.status_updated.connect(self.egg.set_text)
        self.worker.start()

    def on_process_finished(self, success, message):

        if success == True:
            QMessageBox.information(self, "Succès", message)
            self.finish_btn.setVisible(True)
            self.check_progress()
            
        else:
            QMessageBox.critical(self, "Erreur", message)
            self.retry_btn.setVisible(True)
            self.finish_btn.setVisible(True)
        
        self.worker = None

    def on_error(self, error):
        QMessageBox.critical(self, "Erreur lors de la preparation: ", error)
        self.prep_retry_btn.setVisible(True)
        self.worker = None

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "Confirmation", 
                "Une opération est en cours. Voulez-vous vraiment quitter?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.worker.terminate()
                self.worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
    
    def center_window(self):
        screen = QApplication.desktop().screenGeometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )
