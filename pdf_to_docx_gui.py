import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox
)
from PyQt6.QtCore import Qt

class PDFToDocxConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to DOCX Converter")
        self.resize(400, 200)

        self.pdf_files = []
        self.output_dir = ""

        layout = QVBoxLayout()

        self.select_pdf_button = QPushButton("選擇多個 PDF 檔案")
        self.select_pdf_button.clicked.connect(self.select_pdf_files)
        layout.addWidget(self.select_pdf_button)

        self.pdf_label = QLabel("尚未選擇 PDF 檔案")
        self.pdf_label.setWordWrap(True)
        layout.addWidget(self.pdf_label)

        self.select_output_button = QPushButton("選擇轉換後儲存路徑")
        self.select_output_button.clicked.connect(self.select_output_directory)
        layout.addWidget(self.select_output_button)

        self.output_label = QLabel("尚未選擇儲存路徑")
        self.output_label.setWordWrap(True)
        layout.addWidget(self.output_label)

        self.convert_button = QPushButton("開始轉換")
        self.convert_button.clicked.connect(self.convert_files)
        layout.addWidget(self.convert_button)

        self.setLayout(layout)

    def select_pdf_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "選擇 PDF 檔案", "", "PDF Files (*.pdf)"
        )
        if files:
            self.pdf_files = files
            self.pdf_label.setText("\n".join(self.pdf_files))
        else:
            self.pdf_label.setText("尚未選擇 PDF 檔案")

    def select_output_directory(self):
        directory = QFileDialog.getExistingDirectory(
            self, "選擇儲存資料夾", ""
        )
        if directory:
            self.output_dir = directory
            self.output_label.setText(self.output_dir)
        else:
            self.output_label.setText("尚未選擇儲存路徑")

    def convert_files(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "警告", "請先選擇 PDF 檔案")
            return
        if not self.output_dir:
            QMessageBox.warning(self, "警告", "請先選擇儲存路徑")
            return

        from pdf2docx import Converter
        import os
        import zipfile

        docx_files = []
        try:
            for pdf_path in self.pdf_files:
                filename = os.path.basename(pdf_path)
                docx_filename = os.path.splitext(filename)[0] + ".docx"
                docx_path = os.path.join(self.output_dir, docx_filename)

                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()

                docx_files.append(docx_path)

            # 建立 zip 檔案
            zip_path = os.path.join(self.output_dir, "converted_docx_files.zip")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file in docx_files:
                    arcname = os.path.basename(file)
                    zipf.write(file, arcname)

            QMessageBox.information(self, "完成", f"轉換完成，ZIP 檔案已儲存於:\n{zip_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"轉換過程發生錯誤:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToDocxConverter()
    window.show()
    sys.exit(app.exec())
