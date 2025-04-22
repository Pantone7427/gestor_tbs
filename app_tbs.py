import sys
import os
import pandas as pd
import fitz  # PyMuPDF
import traceback
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, 
                           QVBoxLayout, QHBoxLayout, QFileDialog, QProgressBar, 
                           QWidget, QMessageBox, QTextEdit, QGroupBox, QGridLayout)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont, QIcon, QColor

class PDFProcessor(QThread):
    """
    Clase worker que maneja el procesamiento de PDFs en un hilo separado
    para evitar que la interfaz gráfica se congele durante operaciones largas.
    """
    progress_update = pyqtSignal(int)
    log_update = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, tb_path, soporte_path, excel_path, output_folder):
        """
        Inicializa el procesador de PDFs.
        
        Args:
            tb_path (str): Ruta al archivo PDF de TBs
            soporte_path (str): Ruta al archivo PDF de soportes
            excel_path (str): Ruta al archivo Excel
            output_folder (str): Carpeta de salida para los PDFs generados
        """
        super().__init__()
        self.tb_path = tb_path
        self.soporte_path = soporte_path
        self.excel_path = excel_path
        self.output_folder = output_folder
        
    def log(self, message):
        """Emite un mensaje para el registro de la aplicación"""
        self.log_update.emit(message)
        
    def run(self):
        """Método principal que se ejecuta cuando se inicia el hilo"""
        try:
            self.log("Iniciando procesamiento...")
            
            # 1. Leer archivo Excel
            self.log("Leyendo archivo Excel...")
            df = self.read_excel_file()
            if df is None:
                self.finished_signal.emit(False, "Error al leer el archivo Excel")
                return
                
            # 2. Procesar PDF de TBs
            self.log("Extrayendo TBs del PDF...")
            tb_pages = self.extract_tb_pages()
            if tb_pages is None:
                self.finished_signal.emit(False, "Error al procesar el PDF de TBs")
                return
                
            # 3. Procesar PDF de soportes
            self.log("Extrayendo soportes de pago del PDF...")
            soporte_regions = self.extract_soporte_regions()
            if soporte_regions is None:
                self.finished_signal.emit(False, "Error al procesar el PDF de soportes")
                return
                
            # 4. Crear archivos PDF combinados
            self.log("Generando archivos PDF individuales...")
            success = self.create_combined_pdfs(df, tb_pages, soporte_regions)
            if not success:
                self.finished_signal.emit(False, "Error al generar archivos PDF")
                return
                
            self.log("Procesamiento completado con éxito!")
            self.finished_signal.emit(True, f"Se han generado correctamente los archivos PDF en: {self.output_folder}")
            
        except Exception as e:
            error_msg = f"Error durante el procesamiento: {str(e)}\n{traceback.format_exc()}"
            self.log(error_msg)
            self.finished_signal.emit(False, error_msg)
    
    def read_excel_file(self):
        """
        Lee el archivo Excel con la información de las TBs
        
        Returns:
            pandas.DataFrame: DataFrame con la información del Excel, o None si hay error
        """
        try:
            # Leer archivo Excel con las columnas "No Egreso" y "Girado a"
            df = pd.read_excel(self.excel_path)
            
            # Verificar que las columnas requeridas existan
            required_columns = ["No Egreso", "Girado a"]
            for col in required_columns:
                if col not in df.columns:
                    self.log(f"Error: Columna '{col}' no encontrada en el archivo Excel")
                    return None
                    
            self.log(f"Se encontraron {len(df)} registros en el Excel")
            return df
            
        except Exception as e:
            self.log(f"Error al leer el archivo Excel: {str(e)}")
            return None
            
    def extract_tb_pages(self):
        """
        Extrae cada página del PDF de TBs como un documento PDF separado
        
        Returns:
            list: Lista de objetos DocumentPage o None si hay error
        """
        try:
            # Abrir el PDF de TBs
            doc = fitz.open(self.tb_path)
            
            # Lista para almacenar cada página de TB
            tb_pages = []
            
            # Extraer cada página como un documento separado
            total_pages = len(doc)
            for i in range(total_pages):
                # Crear un nuevo documento con esta página
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=i, to_page=i)
                tb_pages.append(new_doc)
                
                # Actualizar progreso
                progress = int((i + 1) / total_pages * 33)  # Primer tercio del proceso
                self.progress_update.emit(progress)
                
            self.log(f"Se extrajeron {len(tb_pages)} TBs del PDF")
            return tb_pages
            
        except Exception as e:
            self.log(f"Error al extraer TBs: {str(e)}")
            return None
    
    def extract_soporte_regions(self):
        """
        Identifica y extrae regiones que contienen soportes de pago del PDF
        
        Returns:
            list: Lista de tuplas (página, región) para cada soporte, o None si hay error
        """
        try:
            # Abrir el PDF de soportes
            doc = fitz.open(self.soporte_path)
            
            # Ajustamos el algoritmo para tener regiones que se superpongan ligeramente
            soporte_regions = []
            total_pages = len(doc)
            
            for page_num in range(total_pages):
                page = doc[page_num]
                page_height = page.rect.height
                page_width = page.rect.width
                
                # En lugar de dividir en tercios exactos, creamos regiones con superposición
                # Esto ayudará a evitar cortar información en los bordes
                
                # Primera región (superior) - un poco más grande
                r1 = fitz.Rect(0, 0, page_width, page_height * 0.34)
                soporte_regions.append((page_num, r1))
                
                # Segunda región (medio) - superposición con primera y tercera
                r2 = fitz.Rect(0, page_height * 0.32, page_width, page_height * 0.68)
                soporte_regions.append((page_num, r2))
                
                # Tercera región (inferior) - empieza más arriba para capturar el encabezado completo
                r3 = fitz.Rect(0, page_height * 0.64, page_width, page_height)
                soporte_regions.append((page_num, r3))
                
                # Actualizar progreso
                progress = 33 + int((page_num + 1) / total_pages * 33)  # Segundo tercio del proceso
                self.progress_update.emit(progress)
                
            self.log(f"Se detectaron {len(soporte_regions)} posibles regiones de soportes")
            return soporte_regions
            
        except Exception as e:
            self.log(f"Error al extraer regiones de soportes: {str(e)}")
            return None
            
    def create_combined_pdfs(self, df, tb_pages, soporte_regions):
        """
        Crea archivos PDF combinados con TB y soporte correspondiente
        
        Args:
            df (pandas.DataFrame): DataFrame con información de TBs
            tb_pages (list): Lista de documentos PDF con páginas de TBs
            soporte_regions (list): Lista de regiones de soportes
            
        Returns:
            bool: True si el proceso fue exitoso, False en caso contrario
        """
        try:
            # Asegurarse de que la carpeta de salida exista
            os.makedirs(self.output_folder, exist_ok=True)
            
            # Por cada registro en el Excel
            total_records = len(df)
            for i, row in df.iterrows():
                # Obtener valores de las columnas
                num_egreso = str(row["No Egreso"])
                girado_a = str(row["Girado a"])
                
                # Verificar que tenemos suficientes TBs y soportes
                if i >= len(tb_pages) or i >= len(soporte_regions):
                    self.log(f"Advertencia: No hay suficientes TBs o soportes para el registro {i+1}")
                    continue
                    
                # Obtener TB correspondiente
                tb_doc = tb_pages[i]
                
                # Obtener soporte correspondiente
                soporte_page_num, region = soporte_regions[i]
                
                # Abrir el PDF de soportes
                soporte_doc = fitz.open(self.soporte_path)
                soporte_page = soporte_doc[soporte_page_num]
                
                # Crear un nuevo documento para el soporte recortado
                cropped_soporte = fitz.open()
                new_page = cropped_soporte.new_page(width=region.width, height=region.height)
                
                # Recortar y copiar la región del soporte
                new_page.show_pdf_page(new_page.rect, soporte_doc, soporte_page_num, clip=region)
                
                # Crear el documento final combinado
                final_doc = fitz.open()
                
                # Agregar la TB
                final_doc.insert_pdf(tb_doc)
                
                # Agregar el soporte recortado
                final_doc.insert_pdf(cropped_soporte)
                
                # Crear nombre de archivo usando No Egreso y Girado a
                # Limpiar el nombre para eliminar caracteres no válidos
                safe_girado_a = "".join(c for c in girado_a if c.isalnum() or c in " ._-").strip()
                filename = f"{num_egreso} - {safe_girado_a}.pdf"
                output_path = os.path.join(self.output_folder, filename)
                
                # Guardar el documento final
                final_doc.save(output_path)
                final_doc.close()
                
                # Cerrar documentos
                cropped_soporte.close()
                
                # Actualizar progreso - Ajustamos para asegurar que llegue a 100%
                progress = 66 + int((i + 1) / total_records * 34)
                self.progress_update.emit(progress)
                self.log(f"Generado archivo: {filename}")
                
            soporte_doc.close()
            
            # Cerrar todos los documentos TB
            for doc in tb_pages:
                doc.close()
                
            # Asegurar que el progreso llegue al 100% al finalizar
            self.progress_update.emit(100)
            
            return True
            
        except Exception as e:
            self.log(f"Error al crear PDFs combinados: {str(e)}")
            return False

class MainWindow(QMainWindow):
    """Ventana principal de la aplicación"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
        # Inicializar variables para almacenar rutas de archivos
        self.tb_path = None
        self.soporte_path = None
        self.excel_path = None
        self.output_folder = None
        
        # Variable para almacenar el worker thread
        self.worker = None
        
    def init_ui(self):
        """Inicializa la interfaz de usuario"""
        self.setWindowTitle("Procesador de TBs y Soportes de Pago")
        self.setMinimumSize(900, 650)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #cccccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                width: 20px;
            }
            QLabel {
                color: #2c3e50;
                font-size: 12px;
            }
            QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 4px;
                background-color: #ffffff;
                font-family: Consolas, monospace;
                padding: 5px;
            }
        """)
        
        # Widget central y layout principal
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Título de la aplicación
        title_label = QLabel("Procesador de TBs y Soportes de Pago")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 15px;")
        main_layout.addWidget(title_label)
        
        # Grupo de selección de archivos
        file_group = QGroupBox("Selección de Archivos")
        file_layout = QGridLayout()
        file_layout.setVerticalSpacing(15)
        file_layout.setHorizontalSpacing(10)
        
        # Archivo de TBs
        self.tb_label = QLabel("Archivo PDF de TBs:")
        self.tb_path_label = QLabel("No seleccionado")
        self.tb_path_label.setStyleSheet("font-style: italic; color: #7f8c8d;")
        self.tb_button = QPushButton("Seleccionar")
        self.tb_button.setIcon(QIcon.fromTheme("document-open"))
        self.tb_button.clicked.connect(self.select_tb_file)
        
        # Archivo de Soportes
        self.soporte_label = QLabel("Archivo PDF de Soportes:")
        self.soporte_path_label = QLabel("No seleccionado")
        self.soporte_path_label.setStyleSheet("font-style: italic; color: #7f8c8d;")
        self.soporte_button = QPushButton("Seleccionar")
        self.soporte_button.setIcon(QIcon.fromTheme("document-open"))
        self.soporte_button.clicked.connect(self.select_soporte_file)
        
        # Archivo Excel
        self.excel_label = QLabel("Archivo Excel:")
        self.excel_path_label = QLabel("No seleccionado")
        self.excel_path_label.setStyleSheet("font-style: italic; color: #7f8c8d;")
        self.excel_button = QPushButton("Seleccionar")
        self.excel_button.setIcon(QIcon.fromTheme("document-open"))
        self.excel_button.clicked.connect(self.select_excel_file)
        
        # Carpeta de salida
        self.output_label = QLabel("Carpeta de Salida:")
        self.output_path_label = QLabel("No seleccionado")
        self.output_path_label.setStyleSheet("font-style: italic; color: #7f8c8d;")
        self.output_button = QPushButton("Seleccionar")
        self.output_button.setIcon(QIcon.fromTheme("folder-open"))
        self.output_button.clicked.connect(self.select_output_folder)
        
        # Agregar al layout
        file_layout.addWidget(self.tb_label, 0, 0)
        file_layout.addWidget(self.tb_path_label, 0, 1)
        file_layout.addWidget(self.tb_button, 0, 2)
        
        file_layout.addWidget(self.soporte_label, 1, 0)
        file_layout.addWidget(self.soporte_path_label, 1, 1)
        file_layout.addWidget(self.soporte_button, 1, 2)
        
        file_layout.addWidget(self.excel_label, 2, 0)
        file_layout.addWidget(self.excel_path_label, 2, 1)
        file_layout.addWidget(self.excel_button, 2, 2)
        
        file_layout.addWidget(self.output_label, 3, 0)
        file_layout.addWidget(self.output_path_label, 3, 1)
        file_layout.addWidget(self.output_button, 3, 2)
        
        file_layout.setColumnStretch(1, 1)  # La columna del medio se estira
        file_group.setLayout(file_layout)
        
        # Grupo de procesamiento
        process_group = QGroupBox("Procesamiento")
        process_layout = QVBoxLayout()
        process_layout.setSpacing(15)
        
        # Barra de progreso con etiqueta
        progress_label = QLabel("Progreso:")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        
        # Botón de proceso en un contenedor centrado
        button_container = QHBoxLayout()
        button_container.addStretch(1)
        self.process_button = QPushButton("Procesar Archivos")
        self.process_button.setMinimumWidth(200)
        self.process_button.setMinimumHeight(40)
        self.process_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        self.process_button.clicked.connect(self.process_files)
        self.process_button.setEnabled(False)
        button_container.addWidget(self.process_button)
        button_container.addStretch(1)
        
        # Área de registro
        self.log_label = QLabel("Registro de actividad:")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("background-color: #f9f9f9; color: #34495e;")
        
        # Agregar al layout
        process_layout.addWidget(progress_label)
        process_layout.addWidget(self.progress_bar)
        process_layout.addLayout(button_container)
        process_layout.addWidget(self.log_label)
        process_layout.addWidget(self.log_text)
        process_group.setLayout(process_layout)
        
        # Agregar grupos al layout principal
        main_layout.addWidget(file_group)
        main_layout.addWidget(process_group)
        
        # Establecer el widget central
        self.setCentralWidget(central_widget)
        
        # Inicializar el log
        self.log("Aplicación iniciada. Por favor, seleccione los archivos para procesar.")
        
    def log(self, message):
        """Agrega un mensaje al área de registro"""
        # Añadir formato de color según el tipo de mensaje
        if "Error" in message:
            formatted_message = f"<span style='color:#e74c3c;'>{message}</span>"
        elif "Advertencia" in message:
            formatted_message = f"<span style='color:#f39c12;'>{message}</span>"
        elif "completado" in message.lower() or "éxito" in message.lower() or "generado" in message.lower():
            formatted_message = f"<span style='color:#27ae60;'>{message}</span>"
        else:
            formatted_message = message
            
        # Añadir timestamp
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        entry = f"[{timestamp}] {formatted_message}"
        
        self.log_text.append(entry)
        # Desplazarse al final
        cursor = self.log_text.textCursor()
        cursor.movePosition(cursor.End)
        self.log_text.setTextCursor(cursor)
        
    def check_files_selected(self):
        """Verifica si todos los archivos necesarios están seleccionados"""
        all_selected = (self.tb_path is not None and 
                       self.soporte_path is not None and 
                       self.excel_path is not None and 
                       self.output_folder is not None)
        self.process_button.setEnabled(all_selected)
        
        # Actualizar estilos visuales para indicar estado completo
        if all_selected:
            self.process_button.setStyleSheet("""
                QPushButton {
                    background-color: #2ecc71;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #27ae60;
                }
            """)
            self.log("Todos los archivos han sido seleccionados. Listo para procesar.")
        
    def select_tb_file(self):
        """Selecciona el archivo PDF de TBs"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Archivo PDF de TBs", "", "Archivos PDF (*.pdf)"
        )
        if file_path:
            self.tb_path = file_path
            self.tb_path_label.setText(os.path.basename(file_path))
            self.tb_path_label.setStyleSheet("color: #2980b9; font-weight: bold;")
            self.log(f"Archivo de TBs seleccionado: {file_path}")
            self.check_files_selected()
            
    def select_soporte_file(self):
        """Selecciona el archivo PDF de soportes de pago"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Archivo PDF de Soportes", "", "Archivos PDF (*.pdf)"
        )
        if file_path:
            self.soporte_path = file_path
            self.soporte_path_label.setText(os.path.basename(file_path))
            self.soporte_path_label.setStyleSheet("color: #2980b9; font-weight: bold;")
            self.log(f"Archivo de Soportes seleccionado: {file_path}")
            self.check_files_selected()
            
    def select_excel_file(self):
        """Selecciona el archivo Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Archivo Excel", "", "Archivos Excel (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path = file_path
            self.excel_path_label.setText(os.path.basename(file_path))
            self.excel_path_label.setStyleSheet("color: #2980b9; font-weight: bold;")
            self.log(f"Archivo Excel seleccionado: {file_path}")
            self.check_files_selected()
            
    def select_output_folder(self):
        """Selecciona la carpeta de salida"""
        folder_path = QFileDialog.getExistingDirectory(
            self, "Seleccionar Carpeta de Salida"
        )
        if folder_path:
            self.output_folder = folder_path
            self.output_path_label.setText(folder_path)
            self.output_path_label.setStyleSheet("color: #2980b9; font-weight: bold;")
            self.log(f"Carpeta de salida seleccionada: {folder_path}")
            self.check_files_selected()
            
    def process_files(self):
        """Inicia el procesamiento de archivos"""
        # Deshabilitar botón durante el proceso
        self.process_button.setEnabled(False)
        self.process_button.setStyleSheet("""
            QPushButton {
                background-color: #cccccc;
                color: #666666;
                font-size: 14px;
            }
        """)
        
        # Restablecer barra de progreso
        self.progress_bar.setValue(0)
        
        # Crear worker thread
        self.worker = PDFProcessor(
            self.tb_path,
            self.soporte_path,
            self.excel_path,
            self.output_folder
        )
        
        # Conectar señales
        self.worker.progress_update.connect(self.update_progress)
        self.worker.log_update.connect(self.log)
        self.worker.finished_signal.connect(self.processing_finished)
        
        # Iniciar procesamiento
        self.log("Iniciando proceso...")
        self.worker.start()
        
    def update_progress(self, value):
        """Actualiza la barra de progreso"""
        self.progress_bar.setValue(value)
        
    def processing_finished(self, success, message):
        """Maneja la finalización del procesamiento"""
        # Asegurar que la barra de progreso esté al 100%
        self.progress_bar.setValue(100)
        
        # Habilitar botón nuevamente
        self.process_button.setEnabled(True)
        self.process_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        
        # Mostrar mensaje de resultado
        if success:
            QMessageBox.information(self, "Proceso Completado", message)
        else:
            QMessageBox.critical(self, "Error", message)
            
        # Limpiar worker
        self.worker = None

def main():
    """Función principal de la aplicación"""
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

