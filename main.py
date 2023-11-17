from PyQt5.QtWidgets import QApplication, QFileDialog, \
    QLabel, QPushButton, QVBoxLayout, QWidget, QMessageBox, \
    QMainWindow, QDesktopWidget, QSizePolicy
from PyQt5.QtGui import QDropEvent
from PyQt5.QtGui import QIcon, QDragEnterEvent
from PyQt5.QtCore import Qt, pyqtSignal, QMimeData
from io import BytesIO
from pptx import Presentation
from PIL import Image
import numpy as np
import pytesseract
import pyttsx3
import cv2
import os

#DEMO1

class DragAndDropLabel(QLabel):
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super(DragAndDropLabel, self).__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setText("Arrastra tu presentación aquí")
        self.setStyleSheet("border: 5px dashed white; color: #ffd166; font-size: 35px;"
                           " font-weight: bold; padding: 40px; margin-top: 20px; background-color: #118ab2; ")

        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        mime_data = event.mimeData()
        if mime_data.hasUrls() and len(mime_data.urls()) == 1:
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        file_path = event.mimeData().urls()[0].toLocalFile()
        self.file_dropped.emit(file_path)

class PowerPointToAudioConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.pptx_path = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        #Configuraciones basicas de la pantalla, titulo de la ventana, icono, color de fondo y geometria de la pantalla.
        pantalla = QDesktopWidget().availableGeometry()
        self.setWindowTitle("Conversión Texto-Audio")
        self.setGeometry(100, 100, 1280, 700)
        self.setWindowIcon(QIcon('esime_original.ico'))
        self.setStyleSheet("background-color: #073b4c;")

        #Titulo del proyecto con su respectivo estilo
        self.titulo1 = QLabel('Convierte un PowerPoint a MP3', self)
        self.titulo1.setStyleSheet("color: #fdf0d5; font-size: 40px; font-weight: bold;")
        layout.addWidget(self.titulo1, alignment=Qt.AlignCenter)

        #Drag and drop y selección de archivos
        self.drag_and_drop_label = DragAndDropLabel(self)
        self.drag_and_drop_label.file_dropped.connect(self.handle_file_dropped)
        layout.addWidget(self.drag_and_drop_label, alignment=Qt.AlignTop)
        height_percentage = 0.6
        label_height = int(pantalla.height() * height_percentage)
        self.drag_and_drop_label.setMinimumHeight(label_height)
        self.drag_and_drop_label.setMaximumHeight(label_height)


        self.boton_subir = QPushButton("Seleccionar Archivo PowerPoint", self)
        self.boton_subir.setStyleSheet("background-color: #202324; color: white; font-size: 20px; font-weight: bold; border: 2px solid white;")
        self.boton_subir.clicked.connect(self.abrir_archivo)
        layout.addWidget(self.boton_subir, alignment=Qt.AlignTop)

        self.boton_convertir = QPushButton("Convertir a Audio", self)
        self.boton_convertir.setStyleSheet("background-color: #202324; color: white; font-size: 20px; font-weight: bold; border: 2px solid white;")
        self.boton_convertir.clicked.connect(self.convertir_a_audio)
        layout.addWidget(self.boton_convertir, alignment=Qt.AlignTop)

        self.etiqueta_seleccionado = QLabel("", self)
        self.etiqueta_seleccionado.setStyleSheet("color: white; font-size: 15px; font-weight: bold;")
        layout.addWidget(self.etiqueta_seleccionado, alignment=Qt.AlignTop)

        self.nombre_archivo_label = QLabel("", self)
        self.nombre_archivo_label.setStyleSheet("color: #2980B9; font-size: 15px; font-weight: bold;")
        layout.addWidget(self.nombre_archivo_label, alignment=Qt.AlignTop)

        self.vista_previa_label = QLabel(self)
        layout.addWidget(self.vista_previa_label, alignment=Qt.AlignTop)

        self.setLayout(layout)

    def handle_file_dropped(self, file_path):
        self.etiqueta_seleccionado.setText("Archivo seleccionado:")
        self.nombre_archivo_label.setText(os.path.basename(file_path))
        self.nombre_archivo_label.setStyleSheet("color: #2980B9; font-size: 15px; font-weight: bold;")
        self.pptx_path = file_path

    def abrir_archivo(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo PowerPoint", "", "Archivos PowerPoint (*.pptx)")
        if archivo:
            nombre_archivo = os.path.basename(archivo)
            self.etiqueta_seleccionado.setText("Archivo seleccionado:")
            self.nombre_archivo_label.setText(nombre_archivo)
            self.nombre_archivo_label.setStyleSheet("color: #2980B9; font-size: 15px; font-weight: bold;")
            self.pptx_path = archivo

    def convertir_a_audio(self):
        if not self.pptx_path:
            QMessageBox.warning(self, "Advertencia", "Primero selecciona un archivo PowerPoint.")
            return

        pptx_info = self.extract_pptx_info_with_ocr(self.pptx_path)
        self.perform_ocr_on_images(pptx_info)

        base_audio_filename = "presentacion"
        self.generate_audio_file(pptx_info, base_audio_filename)

    def extract_table_text(self, table):
        table_text = ""
        for row in table.rows:
            for cell in row.cells:
                table_text += cell.text + " "
            table_text += "\n"
        return table_text

    def extract_pptx_info_with_ocr(self, pptx_path):
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        prs = Presentation(pptx_path)
        pptx_info = {'total_slides': len(prs.slides), 'slides': []}

        for slide in prs.slides:
            slide_info = {'title': "", 'text': "", 'images': [], 'tables': []}

            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = shape.text.strip()
                    if text:
                        if not slide_info['title']:
                            slide_info['title'] = text
                        else:
                            slide_info['text'] += text + "\n"
                elif shape.shape_type == 13:  # Shape type for images
                    img_stream = shape.image.blob
                    img = Image.open(BytesIO(img_stream))
                    slide_info['images'].append({'image': img, 'ocr_text': ''})
                elif shape.has_table:
                    table_text = self.extract_table_text(shape.table)
                    slide_info['tables'].append(table_text)

            pptx_info['slides'].append(slide_info)

        return pptx_info

    def perform_ocr_on_images(self, pptx_info):
        for slide in pptx_info['slides']:
            for img_data in slide['images']:
                img_array = cv2.cvtColor(np.array(img_data['image']), cv2.COLOR_RGB2BGR)
                text = pytesseract.image_to_string(img_array)
                img_data['ocr_text'] = text

    def save_text_to_audio(self, text, base_filename):
        engine = pyttsx3.init()
        engine.setProperty('rate', 150)

        counter = 0
        audio_filename = f"{base_filename}.mp3"

        while os.path.exists(audio_filename):
            counter += 1
            audio_filename = f"{base_filename}_{counter}.mp3"

        engine.save_to_file(text, audio_filename)
        engine.runAndWait()

        QMessageBox.information(self, "Éxito", f"La conversión a audio ha terminado. Se ha guardado en el archivo {os.path.basename(audio_filename)}")

    def generate_audio_file(self, pptx_info, base_filename):
        audio_text = ""

        for i, slide in enumerate(pptx_info['slides']):
            audio_text += f"Diapositiva {i + 1} - Título: {slide['title']}"
            if slide['text']:
                audio_text += f"Contenido: {slide['text']}"
            for j, table_text in enumerate(slide['tables']):
                audio_text += f"Tabla {j + 1} de la diapositiva {i + 1} - Contenido: {table_text}"
            for j, img_data in enumerate(slide['images']):
                if img_data['ocr_text']:
                    audio_text += f"Imagen {j + 1} de la diapositiva {i + 1} - Texto: {img_data['ocr_text']}"
                else:
                    audio_text += f"Imagen {j + 1} de la diapositiva {i + 1} - Sin contenido"

        self.save_text_to_audio(audio_text, base_filename)


if __name__ == "__main__":
    app = QApplication([])
    converter_app = PowerPointToAudioConverter()
    converter_app.show()
    app.exec_()
