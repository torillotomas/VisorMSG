import sys
import os
import tempfile
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QTextEdit, QPushButton, QScrollArea, QFileDialog, QMessageBox)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

class MsgViewer(QMainWindow):
    def __init__(self, msg_file=None):
        super().__init__()
        self.setWindowTitle("Visor de Archivos MSG")
        self.setGeometry(100, 100, 900, 700)
        self.temp_files = []  # Para rastrear archivos temporales
        self.init_ui()
        
        if msg_file:
            self.load_msg_file(msg_file)
    
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        self.subject_label = QLabel("Asunto: ")
        self.sender_label = QLabel("Remitente: ")
        self.to_label = QLabel("Destinatarios: ")
        main_layout.addWidget(self.sender_label)
        main_layout.addWidget(self.to_label)
        main_layout.addWidget(self.subject_label)

        self.body_text = QTextEdit()
        self.body_text.setReadOnly(True)
        main_layout.addWidget(QLabel("Cuerpo del Correo:"))
        main_layout.addWidget(self.body_text, stretch=3)

        self.attachments_area = QScrollArea()
        self.attachments_area.setWidgetResizable(True)
        self.attachments_area.setFixedHeight(150)
        self.attachments_container = QWidget()
        self.attachments_layout = QVBoxLayout(self.attachments_container)
        self.attachments_area.setWidget(self.attachments_container)
        main_layout.addWidget(QLabel("Adjuntos:"))
        main_layout.addWidget(self.attachments_area, stretch=1)

        open_button = QPushButton("Abrir otro archivo MSG")
        open_button.clicked.connect(self.open_msg_file)
        main_layout.addWidget(open_button)

        self.setStyleSheet("""
            QMainWindow { background-color: #fff; }
            QLabel { font-size: 14px; color: #1f4e79; }
            QTextEdit { font-size: 13px; border: 1px solid #d3d3d3; }
            QScrollArea { border: 1px solid #d3d3d3; max-height: 150px; }
            QWidget#attachment_widget { border: 1px solid #e0e0e0; }
            QPushButton { background-color: #0078d4; color: #fff; padding: 6px; font-size: 14px; }
            QPushButton:hover { background-color: #005a9e; }
        """)

    def load_msg_file(self, file_path):
        try:
            # Importar extract_msg solo cuando sea necesario
            import extract_msg
            msg = extract_msg.Message(file_path)
            self.subject_label.setText(f"Asunto: {msg.subject or 'Sin asunto'}")
            self.sender_label.setText(f"Remitente: {msg.sender or 'Desconocido'}")
            self.to_label.setText(f"Destinatarios: {msg.to or 'Sin destinatarios'}")

            html_body = msg.htmlBody
            if html_body and isinstance(html_body, (str, bytes)):
                if isinstance(html_body, bytes):
                    html_body = html_body.decode('utf-8', errors='ignore')
                cid_images = self.extract_cid_images(msg)
                for cid, img_path in cid_images.items():
                    html_body = html_body.replace(f"cid:{cid}", f"file:///{img_path.replace('\\\\', '/').replace('\\', '/')}")
                self.body_text.setHtml(html_body)
            else:
                self.body_text.setPlainText(msg.body or "Sin contenido")

            for i in reversed(range(self.attachments_layout.count())):
                widget = self.attachments_layout.itemAt(i).widget()
                if widget:
                    widget.deleteLater()
            
            self.attachments = msg.attachments
            for attachment in self.attachments:
                self.add_attachment_widget(attachment)
            
            msg.close()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cargar el archivo: {str(e)}")

    def extract_cid_images(self, msg):
        cid_images = {}
        for attachment in msg.attachments:
            cid = getattr(attachment, 'cid', None)
            if cid:
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(attachment.name)[1]) as temp_file:
                    temp_file.write(attachment.data)
                    cid_images[cid] = temp_file.name
                    self.temp_files.append(temp_file.name)
        return cid_images

    def add_attachment_widget(self, attachment):
        attachment_widget = QWidget()
        attachment_widget.setObjectName("attachment_widget")
        attachment_layout = QHBoxLayout()
        attachment_widget.setLayout(attachment_layout)

        is_image = attachment.name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))
        
        if is_image:
            pixmap = QPixmap()
            pixmap.loadFromData(attachment.data)
            if not pixmap.isNull():
                label = QLabel()
                label.setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio))
                attachment_layout.addWidget(label)
            else:
                label = QLabel(f"Adjunto: {attachment.name}")
                attachment_layout.addWidget(label)
        else:
            label = QLabel(f"Adjunto: {attachment.name}")
            attachment_layout.addWidget(label)

        download_button = QPushButton(f"Descargar: {attachment.name}")
        download_button.clicked.connect(lambda: self.download_attachment(attachment))
        attachment_layout.addWidget(download_button)

        self.attachments_layout.addWidget(attachment_widget)

    def open_msg_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Abrir archivo MSG", "", "Archivos MSG (*.msg)")
        if file_path:
            self.load_msg_file(file_path)

    def download_attachment(self, attachment):
        save_path, _ = QFileDialog.getSaveFileName(self, "Guardar adjunto", attachment.name)
        if save_path:
            try:
                attachment.save(custom_path=save_path)
                QMessageBox.information(self, "Éxito", f"Adjunto guardado en: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo guardar el adjunto: {str(e)}")

    def closeEvent(self, event):
        for temp_file in self.temp_files:
            try:
                os.unlink(temp_file)
            except:
                pass
        event.accept()

def main():
    app = QApplication(sys.argv)
    viewer = MsgViewer()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()