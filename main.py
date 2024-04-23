# This Python file uses the following encoding: utf-8
import sys
from PySide6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QLineEdit, QVBoxLayout, QHBoxLayout, QLabel
from createXLS import *
from projetCSVAutomatique import *


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.folder_out = ""
        self.file_template = ""
        self.file_log = ""

        self.initUI()

    def initUI(self):
        self.btn_select_folder_out = QPushButton("Sélectionner un dossier d'arriver pour le nouveau xlsx", self)
        self.btn_select_folder_out.clicked.connect(self.select_folder_out)

        self.btn_select_file_template = QPushButton('Sélectionner le fichier template xlsx', self)
        self.btn_select_file_template.clicked.connect(self.select_file_template)

        self.btn_select_file_log = QPushButton("Sélectionner le log pour l'API google drive", self)
        self.btn_select_file_log.clicked.connect(self.select_file_log)

        #self.btn_select_file_template = QPushButton('Sélectionner le fichier template xlsx', self)
        #self.btn_select_file_template.clicked.connect(self.select_file_template)

        self.btn_start = QPushButton('Upload sur le drive', self)
        self.btn_start.clicked.connect(self.demarrage)

        self.label_name_file = QLabel("Rentrez le nom de votre fichier")
        self.txt_name_file = QLineEdit(self)
        self.layout_label_input1 = QHBoxLayout()
        self.layout_label_input1.addWidget(self.label_name_file)
        self.layout_label_input1.addWidget(self.txt_name_file)

        self.label_folder_id = QLabel("Rentrez si vous voulez l'id du dossier google drive")
        self.txt_folder_id = QLineEdit(self)
        self.layout_label_input2 = QHBoxLayout()
        self.layout_label_input2.addWidget(self.label_folder_id)
        self.layout_label_input2.addWidget(self.txt_folder_id)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.btn_select_folder_out)
        self.layout.addWidget(self.btn_select_file_template)
        self.layout.addWidget(self.btn_select_file_log)
        self.layout.addLayout(self.layout_label_input1)
        self.layout.addLayout(self.layout_label_input2)
        self.layout.addWidget(self.btn_start)

        self.setLayout(self.layout)

        self.setWindowTitle('Mon application')
        self.show()

    def select_folder_out(self):
        self.folder_out = QFileDialog.getExistingDirectory(self, 'Sélectionner un dossier')
        print('Dossier sélectionné :', self.folder_out)

    def select_file_template(self):
        self.file_template, _ = QFileDialog.getOpenFileName(self, 'Sélectionner un fichier')
        print('Fichier sélectionné :', self.file_template)

    def select_file_log(self):
        self.file_log, _ = QFileDialog.getOpenFileName(self, 'Sélectionner un fichier')
        print('Fichier sélectionné :', self.file_log)

    def demarrage(self):
        # code à exécuter lorsque le bouton "Démarrer" est cliqué
        print("Le bouton 'Démarrer' a été cliqué !")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec())
