import sys
import os
import tempfile
import zipfile
import logging

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QMessageBox, QPlainTextEdit, QProgressBar
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from pdf2docx import Converter


def resource_path(relative_path):
    """
    獲取資源的絕對路徑，適用於開發環境和 PyInstaller 打包環境。
    """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class QtLogHandler(logging.Handler):
    """
    自定義的日誌處理器，將日誌訊息發送到 PyQt 的 QPlainTextEdit 控件。
    """
    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit

    def emit(self, record):
        msg = self.format(record)
        self.text_edit.appendPlainText(msg)


class ConverterThread(QThread):
    """
    一個獨立的執行緒，用於在背景執行 PDF 到 DOCX 的轉換和壓縮工作。
    """
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self, pdf_files, output_dir):
        super().__init__()
        self.pdf_files = pdf_files
        self.output_dir = output_dir

    def run(self):
        docx_files = []
        total_files = len(self.pdf_files)
        self.log_signal.emit("開始轉換 PDF 檔案...")

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                for index, pdf_path in enumerate(self.pdf_files):
                    filename = os.path.basename(pdf_path)
                    docx_filename = os.path.splitext(filename)[0] + ".docx"
                    docx_path = os.path.join(temp_dir, docx_filename)

                    self.log_signal.emit(f"正在轉換：{filename}")
                    
                    cv = Converter(pdf_path)
                    cv.convert(docx_path, start=0)
                    cv.close()
                    
                    self.log_signal.emit(f"完成：{docx_filename}")
                    docx_files.append(docx_path)

                    progress = int(((index + 1) / total_files) * 100)
                    self.progress_signal.emit(progress)

                zip_path = os.path.join(self.output_dir, "converted_docx_files.zip")
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file in docx_files:
                        arcname = os.path.basename(file)
                        zipf.write(file, arcname)

            self.finished_signal.emit(zip_path)
        except Exception as e:
            self.error_signal.emit(str(e))


class PDFToDocxConverter(QWidget):
    """
    主應用程式視窗，負責建立 UI 並處理使用者互動。
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to DOCX Converter")
        self.resize(500, 450)

        self.pdf_files = []
        self.output_dir = ""
        self.setWindowIcon(QIcon(resource_path("icon.png"))) # 在 __init__ 方法中設置圖示
        self._init_ui()
        self._init_logging()
        self._set_stylesheet()


    def _init_ui(self):
        """
        初始化並設置應用程式的 UI 介面。
        """
        layout = QVBoxLayout()

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

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

    def _init_logging(self):
        """
        初始化並設置日誌系統。
        """
        log_handler = QtLogHandler(self.log_output)
        log_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
        pdf2docx_logger = logging.getLogger("pdf2docx")
        pdf2docx_logger.addHandler(log_handler)
        pdf2docx_logger.setLevel(logging.INFO)

    def _set_stylesheet(self):
        """
        設定整個應用程式的樣式表。
        """
        self.setStyleSheet("""
            QWidget {
                background-color: #2e2e2e;
                color: #ffffff;
                font-family: 'Noto Sans TC', Arial, sans-serif;
            }

            QPushButton {
                background-color: #555555;
                border-radius: 8px;
                padding: 10px;
                color: #ffffff;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #4CAF50; /* 滑鼠懸停時的顏色 */
            }
            QPushButton:pressed {
                background-color: #409040; /* 按下時的顏色 */
            }

            QPlainTextEdit {
                background-color: #1e1e1e;
                border: 1px solid #444444;
                border-radius: 8px;
                padding: 5px;
            }

            QLabel {
                padding: 5px;
            }

            QProgressBar {
                border: 1px solid #444444;
                border-radius: 8px;
                text-align: center;
                height: 25px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 8px;
            }
        """)

    def log(self, message):
        self.log_output.appendPlainText(message)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

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

        self.progress_bar.setValue(0)
        self.thread = ConverterThread(self.pdf_files, self.output_dir)
        self.thread.log_signal.connect(self.log)
        self.thread.progress_signal.connect(self.update_progress)
        self.thread.finished_signal.connect(self.show_success)
        self.thread.error_signal.connect(self.show_error)
        self.thread.start()

    def show_success(self, zip_path):
        QMessageBox.information(self, "完成", f"已儲存於:\n{zip_path}")
        self.progress_bar.setValue(100)

    def show_error(self, error_msg):
        self.log(f"[ERROR] {error_msg}")
        QMessageBox.critical(self, "錯誤", f"轉換過程發生錯誤:\n{error_msg}")
        self.progress_bar.setValue(0)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    window = PDFToDocxConverter()
    window.show()
    sys.exit(app.exec())
