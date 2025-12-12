from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox
import os
import pandas as pd
from file_processor import gan_id_word, gan_id_ppt
from email_utils import send_email
from id_checker import check_id_word, check_id_ppt

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ứng dụng gán ID và gửi email")
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.label = QLabel("Chọn file Excel và thư mục tài liệu:")
        self.btn_excel = QPushButton("Tải file Excel (ID + Gmail)")
        self.btn_folder = QPushButton("Chọn thư mục chứa file Word/PPT")
        self.btn_start = QPushButton("Bắt đầu gán ID và gửi mail")
        self.btn_check = QPushButton("Kiểm tra ID trong file")

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.btn_excel)
        self.layout.addWidget(self.btn_folder)
        self.layout.addWidget(self.btn_start)
        self.layout.addWidget(self.btn_check)

        self.btn_excel.clicked.connect(self.load_excel)
        self.btn_folder.clicked.connect(self.load_folder)
        self.btn_start.clicked.connect(self.process)
        self.btn_check.clicked.connect(self.check_id)

        self.excel_path = ''
        self.folder_path = ''
        self.smtp_config = {
            "from": "meoomimi.03@gmail.com",
            "password": "qntv rqfs bapl bqke",
            "server": "smtp.gmail.com",
            "port": 465
        }

    def load_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Chọn file Excel", "", "*.xlsx")
        if file:
            self.excel_path = file
            self.label.setText(f"Đã chọn Excel: {file}")

    def load_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục chứa file Word/PPT")
        if folder:
            self.folder_path = folder
            self.label.setText(f"Đã chọn thư mục: {folder}")

    def process(self):
        if not self.excel_path or not self.folder_path:
            QMessageBox.warning(self, "Thiếu dữ liệu", "Vui lòng chọn file Excel và thư mục tài liệu.")
            return

        df = pd.read_excel(self.excel_path)
        for idx, row in df.iterrows():
            id_text = str(row['Tên ID'])
            gmails = [g.strip() for g in str(row['Gmail']).split(',')]
            for file in os.listdir(self.folder_path):
                path = os.path.join(self.folder_path, file)
                if file.endswith(".docx"):
                    new_file = gan_id_word(path, id_text, self.folder_path)
                elif file.endswith(".pptx"):
                    new_file = gan_id_ppt(path, id_text, self.folder_path)
                else:
                    continue
                for gmail in gmails:
                    send_email(gmail, new_file, id_text, self.smtp_config)
        QMessageBox.information(self, "Hoàn thành", "Đã gán ID và gửi email thành công!")

    def check_id(self):
        file, _ = QFileDialog.getOpenFileName(self, "Chọn file Word hoặc PPT", "", "Word/PPT Files (*.docx *.pptx)")
        if not file:
            return
        if file.endswith(".docx"):
            id_found = check_id_word(file)
        else:
            id_found = check_id_ppt(file)
        if id_found:
            QMessageBox.information(self, "ID tìm thấy", f"File chứa: {id_found}")
        else:
            QMessageBox.information(self, "Không có ID", "Không tìm thấy ID trong file.")
