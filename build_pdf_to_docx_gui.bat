@echo off
REM 使用 pyinstaller 打包 pdf_to_docx_gui.py 成單一 exe 檔案
pyinstaller --onefile pdf_to_docx_gui.py

echo 打包完成，請查看 dist 資料夾
pause
