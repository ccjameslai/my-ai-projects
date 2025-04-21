這是一個強大的 Python 工具，可以將多種文件格式（PDF、Word、Excel、CSV、JPG 等）自動解析為 Markdown 格式，並支援圖片擷取。  
✨ 本工具結合 Azure Document Intelligence（原 Azure Form Recognizer）進行文字與圖像智慧解析，適合進行文件轉換、資訊擷取、報告生成等用途。

---

## 🚀 支援的功能

✅ PDF 文件 → Markdown + 圖片提取
✅ Word (.docx) 文件 → Markdown
✅ Excel (.xlsx) 表格 → Markdown 表格
✅ Excel 圖表 → PNG 圖片 + Markdown 插入
✅ CSV → Markdown 表格
✅ 圖片（JPG、PNG）→ Azure AI 解讀並轉 Markdown
✅ 圖片與 `<figure>` 位置一一對應
✅ 自動建立本地資料夾：`md_files/`, `image_files/`, `split_files/`

---

## 🧠 使用到的 Azure AI 技術

- [Azure Document Intelligence](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/)
  - 模型：`prebuilt-layout` 用於偵測文字、表格與圖片區塊

---

## 🔧 安裝方式

### 安裝必要套件（pip）：

```bash
pip install pandas openpyxl python-docx PyMuPDF Pillow markdownify xlwings azure-ai-documentintelligence langchain pywin32
```

或使用：

```bash
pip install -r requirements.txt
```

---

## 📦 使用方法

```bash
python azure_doc2markdown.py ./path/to/your/file.pdf
```

支援的檔案副檔名：

- `.pdf`, `.docx`, `.xlsx`, `.csv`, `.jpg`, `.jpeg`, `.png`

---

## 📁 輸出資料夾說明

| 資料夾名稱     | 說明                             |
| -------------- | -------------------------------- |
| `md_files/`    | 儲存轉換後的 Markdown 檔案       |
| `image_files/` | 儲存擷取的圖表與圖片（PNG）      |
| `split_files/` | PDF 拆頁處理中的臨時檔案儲存位置 |

---

## ⚠️ 注意事項

- 本工具僅支援 **Windows 作業系統**（需用到 `xlwings`, `win32com`）
- 請事先註冊 Azure，並取得以下環境變數：

```bash
AZURE_AIDOC_ENDPOINT=https://your-formrecognizer.cognitiveservices.azure.com/
AZURE_AIDOC_KEY=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

建議放入 `.env` 或透過系統環境變數設定。

---

## 📄 範例輸出 (Markdown 格式)

```markdown
# 工作表：營收報告

| 月份 | 營收  | 成本  |
|------|-------|-------|
| 1月  | 12000 | 5000  |
| 2月  | 15000 | 6000  |

![Image](../image_files/Sheet1_chart_1.png)
```

---

## 📚 可拓展功能（TODO）

- [ ] 支援 PDF 多欄位自動偵測
- [ ] 整合輸出 HTML/PDF 報告
- [ ] 支援圖片 OCR 轉文字（使用 Azure OCR 模型）
- [ ] 多語系 Markdown 排版支援
- [ ] GUI 圖形介面支援

---

## 👨‍💻 開發者資訊

作者：James Lai 😎  
協力 AI：ChatGPT + Azure AI