from __future__ import annotations

import os
import threading
from typing import Callable

from openpyxl.utils import column_index_from_string

from .browser import BrowserSession
from .debug import dlog
from .excel_ops import scan_rows_to_process, write_xlsx_from_txts


# ──────────────────────────────────────────────────────────────────────────────
# Phase 1: 抓網頁內容 → 存 txt
# ──────────────────────────────────────────────────────────────────────────────

class ExcelProcessor(threading.Thread):
    def __init__(
        self,
        filepath: str,
        progress_cb: Callable[[int, str, int, int], None],
        status_cb: Callable[[str], None],
        content_cb: Callable[[str], None] | None = None,
        error_cb: Callable[[int, str, str], None] | None = None,
        start_row: int = 2,
        url_col: str = "AV",
        single_row: bool = False,
    ):
        super().__init__(daemon=True)
        self._filepath    = filepath
        self._progress_cb = progress_cb
        self._status_cb   = status_cb
        self._content_cb  = content_cb
        self._error_cb    = error_cb
        self._start_row   = start_row
        self._url_col     = url_col
        self._single_row  = single_row
        self._dest_col    = column_index_from_string(url_col.upper().strip()) + 1
        self._stop_event  = threading.Event()

    def stop(self) -> None:
        self._stop_event.set()

    def _txt_path(self, row_num: int) -> str:
        """Return the expected txt path for a given row: {base}/{row//100}/{row}.txt"""
        base_dir = os.path.splitext(self._filepath)[0]
        return os.path.join(base_dir, str(row_num // 100), f"{row_num}.txt")

    def _save_to_file(self, row_num: int, text: str) -> str:
        """Save text to the deterministic txt path; returns the file path."""
        file_path = self._txt_path(row_num)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(text or "")
        dlog(f"_save_to_file: {file_path} ({len(text or '')} 字元)")
        return file_path

    def run(self) -> None:
        dlog(f"ExcelProcessor.run: 開始 filepath={self._filepath} "
             f"start_row={self._start_row} url_col={self._url_col} "
             f"single_row={self._single_row}")
        try:
            if self._single_row and self._start_row == 1:
                dlog("ExcelProcessor.run: 單列模式 start_row=1，為標題列，結束")
                self._status_cb("此為標題列")
                return

            dlog("ExcelProcessor.run: 呼叫 scan_rows_to_process ...")
            rows = scan_rows_to_process(
                self._filepath,
                self._start_row,
                self._url_col,
                single_row=self._single_row,
            )
            dlog(f"ExcelProcessor.run: 掃描完成，{len(rows)} 列")

            if not rows:
                dlog("ExcelProcessor.run: 無需處理，結束")
                self._status_cb("沒有需要處理的列")
                return

            total = len(rows)
            saved = 0
            current = 0
            dlog("ExcelProcessor.run: 啟動 BrowserSession ...")
            with BrowserSession(self._filepath) as session:
                dlog("ExcelProcessor.run: 進入處理迴圈")
                for row_num, url in rows:
                    if self._stop_event.is_set():
                        dlog("ExcelProcessor.run: 收到停止訊號，跳出迴圈")
                        break

                    # 非單列模式：txt 存在且 size > 0 → 跳過
                    if not self._single_row:
                        p = self._txt_path(row_num)
                        if os.path.isfile(p) and os.path.getsize(p) > 0:
                            dlog(f"ExcelProcessor.run: 第 {row_num} 列已有 txt，跳過")
                            continue

                    current += 1
                    dlog(f"ExcelProcessor.run: ── 處理第 {row_num} 列 ({current}/{total}) ──")
                    self._progress_cb(row_num, url, current, total)
                    text = None
                    fetch_ok = False
                    for attempt in range(1, 4):
                        dlog(f"ExcelProcessor.run: 第 {row_num} 列，嘗試第 {attempt} 次擷取")
                        try:
                            text = session.fetch_text(url)
                            if not text or not text.strip():
                                raise ValueError("擷取內容為空")
                            dlog(f"ExcelProcessor.run: 第 {row_num} 列，第 {attempt} 次成功，{len(text)} 字元")
                            fetch_ok = True
                            break
                        except Exception as e:
                            dlog(f"ExcelProcessor.run: 第 {row_num} 列，第 {attempt} 次失敗：{e}")
                            if attempt == 3:
                                dlog(f"ExcelProcessor.run: 第 {row_num} 列，三次皆失敗，停止")
                                if self._error_cb:
                                    self._error_cb(row_num, url, str(e))
                                self._stop_event.set()

                    if not fetch_ok:
                        break  # 離開主迴圈

                    if self._content_cb:
                        self._content_cb(text)
                    self._save_to_file(row_num, text)
                    saved += 1
                    dlog(f"ExcelProcessor.run: 第 {row_num} 列，txt 已存")

            if self._stop_event.is_set():
                self._status_cb(f"已停止（已存 {saved} 列至 txt）")
            else:
                self._status_cb(f"擷取完成（共 {saved} 列已存至 txt）")

        except Exception as e:
            dlog(f"ExcelProcessor.run: 未預期例外 {type(e).__name__}: {e}")
            self._status_cb(f"錯誤：{e}")


# ──────────────────────────────────────────────────────────────────────────────
# Phase 2: 掃描 txt 資料夾 → 寫入 xlsx
# ──────────────────────────────────────────────────────────────────────────────

def _scan_txt_folder(filepath: str) -> list[tuple[int, str]]:
    """Walk the folder named after the xlsx file and collect (row_num, path) pairs."""
    base_dir = os.path.splitext(os.path.abspath(filepath))[0]
    if not os.path.isdir(base_dir):
        return []
    pairs: list[tuple[int, str]] = []
    for dirpath, _, filenames in os.walk(base_dir):
        for name in filenames:
            if not name.endswith(".txt"):
                continue
            stem = name[:-4]
            if stem.isdigit():
                pairs.append((int(stem), os.path.join(dirpath, name)))
    pairs.sort(key=lambda x: x[0])
    dlog(f"_scan_txt_folder: 找到 {len(pairs)} 個 txt 檔")
    return pairs


class TxtFolderFiller(threading.Thread):
    """Reads all txt files from the xlsx's companion folder and writes them to xlsx via COM."""

    def __init__(
        self,
        filepath: str,
        url_col: str,
        status_cb: Callable[[str], None],
        start_row: int = 2,
        log_cb=None,
        single_row: bool = False,
    ):
        super().__init__(daemon=True)
        self._filepath    = filepath
        self._url_col_idx = column_index_from_string(url_col.upper().strip())
        self._status_cb   = status_cb
        self._start_row   = start_row
        self._log_cb      = log_cb
        self._single_row  = single_row
        self._stop_event  = threading.Event()

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        dlog(f"TxtFolderFiller.run: filepath={self._filepath} url_col_idx={self._url_col_idx} start_row={self._start_row} single_row={self._single_row}")
        try:
            pairs = _scan_txt_folder(self._filepath)
            if self._single_row:
                pairs = [(r, p) for r, p in pairs if r == self._start_row]
            elif self._start_row > 2:
                pairs = [(r, p) for r, p in pairs if r >= self._start_row]
            if not pairs:
                self._status_cb("找不到 txt 檔（請先執行擷取）")
                return
            write_xlsx_from_txts(
                self._filepath,
                pairs,
                self._url_col_idx,
                status_cb=self._status_cb,
                stop_event=self._stop_event,
                log_cb=self._log_cb,
            )
        except Exception as e:
            dlog(f"TxtFolderFiller.run: 例外 {type(e).__name__}: {e}")
            if self._log_cb:
                self._log_cb(f"錯誤：{e}\n")
            self._status_cb(f"錯誤：{e}")
