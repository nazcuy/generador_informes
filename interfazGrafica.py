import sys
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QMessageBox, 
    QFileDialog, QVBoxLayout, QWidget, QTextEdit,
    QProgressBar, QPushButton, QLabel)
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtCore import Qt

class GeneradorPDFapp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generador Informes obras COMPLETAR")
        self.setGeometry(100, 100, 700, 600) # x, y, ancho, alto

        # Fuente personalizada
        font_path = os.path.join(os.path.dirname(__file__), "fonts", "EncodeSans-VariableFont_wdth,wght.ttf.ttf")
        font_id = QFontDatabase.addApplicationFont(font_path)
        font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        font = QFont(font_family, 10)
        self.setFont(font)

        # Widget central
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
##########
        # Layout principal
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.central_widget)

        # Área de texto para mostrar logs
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

        # Barra de progreso
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Botones
        self.select_button = QPushButton("Seleccionar Carpeta de Imágenes")
        self.select_button.clicked.connect(self.seleccionar_carpeta)
        layout.addWidget(self.select_button)

        self.generate_button = QPushButton("Generar Informe PDF")
        self.generate_button.clicked.connect(self.generar_informe)
        layout.addWidget(self.generate_button)

        # Etiqueta para mostrar la carpeta seleccionada
        self.selected_folder_label = QLabel("Carpeta seleccionada: Ninguna")
        layout.addWidget(self.selected_folder_label)

    def log(self, message):
        self.log_area.append(message)
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())

    def seleccionar_carpeta(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta de Imágenes", os.getcwd())
        if folder:
            self.selected_folder = folder
            self.selected_folder_label.setText(f"Carpeta seleccionada: {folder}")
            self.log(f"Carpeta seleccionada: {folder}")
        else:
            self.log("Selección de carpeta cancelada.")

    def generar_informe(self):
        if not hasattr(self, 'selected_folder'):
            QMessageBox.warning(self, "Advertencia", "Por favor, seleccione una carpeta de imágenes primero.")
            return

        self.log("Iniciando generación del informe PDF...")
        
        # Aquí iría la lógica para generar el informe PDF usando la carpeta seleccionada.
        # Por ahora, solo simularemos el proceso con un bucle.
        
        total_steps = 10  # Simulamos 10 pasos en el proceso