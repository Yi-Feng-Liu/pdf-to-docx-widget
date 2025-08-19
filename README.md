# PDF to DOCX Converter with PyQt6

這是一個使用 PyQt6 製作的簡易工具，可以將多個 PDF 檔案轉換成 DOCX 格式，並將所有轉換後的 DOCX 檔案打包成 ZIP 檔案。

## 功能

- 選擇多個 PDF 檔案
- 選擇轉換後的儲存路徑
- 轉換後的 DOCX 檔名與 PDF 檔名相同，副檔名改為 `.docx`
- 將所有轉換後的 DOCX 檔案打包成 ZIP 檔案輸出

## 使用說明

1. 確認已安裝 Python 3.8以上
2. 安裝必要套件：
   ```
   pip install PyQt6 pdf2docx
   ```
3. 執行程式：
   ```
   python pdf_to_docx_gui.py
   ```
4. 在介面中選擇多個 PDF 檔案，選擇轉換後的儲存路徑，點擊「開始轉換」即可。

## 打包成執行檔

已提供 `build_pdf_to_docx_gui.bat` 批次檔，使用 PyInstaller 打包：

```
.\build_pdf_to_docx_gui.bat
```

打包完成後，執行檔會在 `dist` 資料夾中。

## 注意事項

- 轉換過程中可能因 PDF 格式複雜而導致轉換結果不理想，請視需求調整或使用其他工具。
- 請確保環境中已安裝 PyQt6 與 pdf2docx。
