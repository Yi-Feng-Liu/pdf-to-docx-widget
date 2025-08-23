@echo off
echo building "PDF Converter" ...
pyinstaller --onefile --windowed --icon=icon.ico --name "PDF Converter" --add-data "icon.png;." pdf_to_docx_gui.py

echo Finished
