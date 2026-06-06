# copyWebContent2Excel — 實作計畫

## Context
從 Excel 的 AV 欄逐列讀取專利網址（justia.com），用 Playwright 驅動 Chrome 開啟後擷取頁面主體全文，貼入同列 AW 欄，直到 AV 欄無資料為止。列高鎖定單行，不隨內容擴展。超過 Excel 單格上限（32,767 字）時截斷。UI 顯示進度，處理邏輯在背景執行緒執行。最終用 PyInstaller 打包成 .exe。

---

## 環境安裝（套件記錄）

### Python 版本
Python 3.13（穩定版）

### 安裝指令
```bat
pip install --upgrade pip
pip install playwright openpyxl pyinstaller
playwright install chromium
```

> 備註：`playwright install chromium` 下載 Playwright 內建 Chromium（約 150MB）。
> 若想用已安裝的 Chrome，可改用 `channel='chrome'`（不需下載）。

### requirements.txt
```
playwright>=1.40
openpyxl>=3.1
pyinstaller>=6.0
```

---

## 欄位對應

| 用途 | 欄位字母 | openpyxl 欄號 |
|------|----------|--------------|
| URL 來源 | AV | 48 |
| 擷取內容目的地 | AW | 49 |

`column_index_from_string('AV')` → 48

---

## 檔案結構

```
d:\workspace\copyWebContent2Excel\
├── main.py              # 入口點：只負責啟動 UI
├── requirements.txt     # 套件清單
├── build.bat            # PyInstaller 打包指令
└── sub/
    ├── __init__.py      # 空白，標記為 package
    ├── config.py        # 常數：欄號、字元上限、列高
    ├── browser.py       # Playwright 開啟瀏覽器、擷取文字
    ├── excel_ops.py     # openpyxl 讀取 URL、寫入結果、設列高
    ├── processor.py     # 執行緒：協調 browser + excel_ops
    └── ui.py            # tkinter UI（視窗、按鈕、進度顯示）
```

---

## 各檔案職責

### main.py（入口點）
```python
from sub.ui import App
if __name__ == '__main__':
    App().run()
```

### sub/config.py
```python
from openpyxl.utils import column_index_from_string
URL_COL    = column_index_from_string('AV')  # 48
DEST_COL   = URL_COL + 1                      # 49 = AW
MAX_CHARS  = 32767
ROW_HEIGHT = 15
```

### sub/browser.py
- `BrowserSession` 類別，用 `with` 語法管理 Playwright 生命週期
- `fetch_text(url) → str`：`page.goto` + `page.inner_text('body')`
- 整個處理批次共用同一個 browser / page，不每次重開

### sub/excel_ops.py
- `load_urls(filepath) → list[(row_num, url)]`：讀取所有非空 AV 欄
- `write_result(ws, row_num, text)`：寫入 AW 欄並設列高為 `ROW_HEIGHT`
- `save(wb, filepath)`：存檔

### sub/processor.py
- `ExcelProcessor(threading.Thread)`
  - `__init__(filepath, progress_cb, status_cb)`
  - `run()`：使用 `BrowserSession` + `excel_ops` 逐列處理，每列後存檔
  - `stop()`：設 `_stop` flag，下一列前檢查

### sub/ui.py（tkinter）
```
+--------------------------------------------------+
|  [選擇 Excel 檔案]  path/to/file.xlsx            |
|  [開始]  [停止]                                  |
|  進度：第 3 列 / 共 25 筆                        |
|  目前網址：https://patents.justia.com/...        |
|  狀態：處理中...                                 |
+--------------------------------------------------+
```
- UI 更新透過 `root.after(0, callback)` 保持執行緒安全
- 開始時禁用「選擇檔案」與「開始」；完成/停止後恢復
- `App.run()` 呼叫 `root.mainloop()`

### build.bat
```bat
pyinstaller --onefile --windowed --name copyWebContent2Excel main.py
pause
```

---

## 執行步驟

1. 安裝 Python 3.13，加入 PATH
2. `pip install --upgrade pip`
3. `pip install playwright openpyxl pyinstaller`
4. `playwright install chromium`
5. `python main.py` 測試功能
6. 確認 AV 欄 URL 可正確擷取後，執行 `build.bat` 打包

---

## 驗證方式

1. 準備一份 Excel，AV 欄放 2–3 筆 justia 專利網址
2. 執行 `python main.py`，選取該 Excel，按開始
3. 確認 Chromium 自動開啟各網址
4. 確認 AW 欄寫入全文（非空）
5. 確認各列高度為單行（15 points）
6. 確認超長內容截斷在 32,767 字元
