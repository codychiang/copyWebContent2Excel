# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 執行與打包

```batch
# 從原始碼執行
python main.py
# 或
run.bat          # 使用 .venv\Scripts\python.exe

# 打包成獨立 .exe
build.bat        # pyinstaller --onedir --windowed → dist/copyWebContent2Excel/
```

無自動化測試，直接執行 app 並選取 Excel 檔案進行手動驗證。

## 架構

兩階段流程，用於擷取專利頁面（justia.com）並寫入 Excel：

**第一階段 — 抓取**（`sub/processor.py` 的 `ExcelProcessor` 執行緒）
1. `scan_rows_to_process()` 透過 openpyxl 唯讀讀取 Excel 中的網址
2. `BrowserSession`（Playwright 持久化 Chrome）逐列擷取頁面
3. 內容存至 `{xlsx_base}/{row//100}/{row}.txt`
4. 跳過邏輯：txt 檔存在且 size > 0 → 該列已處理完畢

**第二階段 — 寫入**（`sub/processor.py` 的 `TxtFolderFiller` 執行緒）
1. 掃描 txt 資料夾，收集所有 `{列號}.txt`
2. 透過 `win32com.client.DispatchEx("Excel.Application")` 寫入 Excel（必須用 `DispatchEx`，不可用 `Dispatch`，否則會連到已開啟的 Excel 實例）
3. COM 執行緒中必須呼叫 `pythoncom.CoInitialize()` / `CoUninitialize()`

## 欄位配置（預設 url_col = AV = index 48）

| 偏移 | 內容 |
|------|------|
| +0 | 網址（不修改） |
| +3 | 超連結 → txt 檔（相對路徑，顯示文字為純檔名） |
| +4 | 字元數（字串格式） |
| +5 | 占用格數 — 在內容寫完後回填，確保回推切割的格數正確 |
| +6+ | 各 chunk 內容（每格 ≤32000 字） |

## 重要實作細節

**文字切割**（`sub/excel_ops.py` 的 `split_text`）：在 32000 字限制前找最靠右的斷點（`\n .。;；!！?？`），找不到則硬切。

**超連結 COM 呼叫**必須用位置參數（late-bound dispatch 會忽略關鍵字參數）：
```python
ws.Hyperlinks.Add(anchor_cell, address, "", "", display_text)
```

**儲存格樣式 reset**：寫入字元數 / 格數前須呼叫 `cell.Style = "Normal"`，清除前次執行殘留的 Hyperlink 樣式（藍字＋底線）。

**網址 fallback**：AV 欄為空時，從 K 欄（核准號）或 M 欄（公開號）組成 Justia 網址。

**Chrome profile**：存於 `{xlsx_base}_chrome_profile`，每個 xlsx 獨立一份，兩個 app 實例同時執行不會互相干擾。

**警報聲**：擷取失敗時以 `winsound.PlaySound` + `SND_LOOP | SND_ASYNC` 持續播放；`SND_PURGE` 停止。警報只由使用者動作停止，status callback 不會停止它。

**UI 執行緒安全**：worker 執行緒的所有 callback 必須透過 `root.after(0, fn)` 傳回主執行緒。

**按鈕狀態**：由 `_is_crawling` / `_is_filling` 旗標管理（不用 `thread.is_alive()`，因為有時機問題）。只有收到終止訊息才清除旗標（包含：完成／已停止／錯誤／找不到／此為標題列／沒有需要）。

## 模組對照

| 檔案 | 職責 |
|------|------|
| `main.py` | 進入點 |
| `sub/ui.py` | Tkinter GUI、按鈕狀態管理、警報 |
| `sub/processor.py` | `ExcelProcessor` + `TxtFolderFiller` 執行緒 |
| `sub/excel_ops.py` | `scan_rows_to_process`、`split_text`、`ExcelComWriter`、`write_xlsx_from_txts` |
| `sub/browser.py` | Playwright session、Cloudflare 偵測、文字修剪 |
| `sub/session.py` | `setting.ini` 讀寫（上次 filepath、start_row、url_col） |
| `sub/config.py` | 欄位索引常數、字元上限、列高 |
| `sub/debug.py` | `dlog()` 帶時間戳記的日誌（切換 `DEBUG` 旗標啟用） |
