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

        # Layout principal
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setAlignment(Qt.AlignTop)
        self.layout.setSpacing(15)
        self.layout.setContentsMargins(20, 20, 20, 20)
        
        # Variables para almacenar datos
        self.excel_path = ""
        
        # Llamar a métodos para crear la interfaz
        self.crear_interfaz()
        
    def crear_interfaz(self):
        # Etiqueta de título
        self.titulo_label = QLabel("Generador de Informes de Obras COMPLETAR")
        self.titulo_label.setAlignment(Qt.AlignCenter)
        self.titulo_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.layout.addWidget(self.titulo_label)
        
        # Instrucciones
        instrucciones = QLabel(
            "Esta aplicación genera informes PDF a partir de un archivo Excel con datos de obras.\n\n"
            "Pasos a seguir:\n"
            "1. Buscar archivo de obras 'pdf_generator_3000' en la carpeta Z\n"
            "2. Seleccione el archivo usando el botón 'Seleccionar Excel'\n"
            "3. Haga clic en 'Generar PDFs' para crear los informes\n"
            "4. Los PDFs se guardarán en la carpeta 'informes'"
        )
        instrucciones.setStyleSheet("color: #34495e;")
        self.layout.addWidget(instrucciones)
##########
        # Botón para seleccionar archivo Excel
        self.seleccionar_excel_btn = QPushButton("Seleccionar Excel")
        self.seleccionar_excel_btn.clicked.connect(self.seleccionar_archivo_excel)
        self.layout.addWidget(self.seleccionar_excel_btn)

        # Área de texto para mostrar el contenido del Excel
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.layout.addWidget(self.text_edit)

        # Barra de progreso
        self.progreso_bar = QProgressBar()
        self.progreso_bar.setRange(0, 100)
        self.layout.addWidget(self.progreso_bar)

        # Botón para generar PDF
        self.generar_pdf_btn = QPushButton("Generar PDF")
        self.generar_pdf_btn.clicked.connect(self.generar_pdf)
        self.layout.addWidget(self.generar_pdf_btn)