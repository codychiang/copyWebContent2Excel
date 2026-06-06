from __future__ import annotations

import threading
from typing import Callable

from openpyxl.utils import column_index_from_string

from .browser import BrowserSession
from .excel_ops import load_workbook, load_urls_from_file, write_result, save


class ExcelProcessor(threading.Thread):
    def __init__(
        self,
        filepath: str,
        progress_cb: Callable[[int, str], None],
        status_cb: Callable[[str], None],
        start_row: int = 2,
        url_col: str = "AV",
        single_row: bool = False,
    ):
        super().__init__(daemon=True)
        self._filepath   = filepath
        self._progress_cb = progress_cb
        self._status_cb  = status_cb
        self._start_row  = start_row
        self._url_col    = url_col
        self._single_row = single_row
        self._dest_col   = column_index_from_string(url_col.upper().strip()) + 1
        self._stop_event = threading.Event()

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        try:
            rows = load_urls_from_file(
                self._filepath,
                self._start_row,
                self._url_col,
                single_row=self._single_row,
                skip_done=not self._single_row,
            )
            wb, ws = load_workbook(self._filepath)

            if not rows:
                self._status_cb("沒有需要處理的列")
                return

            with BrowserSession() as session:
                for row_num, url in rows:
                    if self._stop_event.is_set():
                        break
                    self._progress_cb(row_num, url)
                    text = None
                    for attempt in range(1, 4):
                        try:
                            text = session.fetch_text(url)
                            break
                        except Exception as e:
                            if attempt == 3:
                                text = f"[擷取失敗] {e}"
                    write_result(ws, row_num, text, self._dest_col)
                    save(wb, self._filepath)

            if self._stop_event.is_set():
                self._status_cb("已停止")
            else:
                self._status_cb("完成")
        except Exception as e:
            self._status_cb(f"錯誤：{e}")
