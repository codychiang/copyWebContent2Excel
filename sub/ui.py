from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox

from .processor import ExcelProcessor
from .session import load_session, save_session


class App:
    def __init__(self):
        self._root = tk.Tk()
        self._root.title("copyWebContent2Excel")
        self._root.resizable(False, False)

        self._filepath  = tk.StringVar(value="（尚未選擇檔案）")
        self._start_row = tk.IntVar(value=2)
        self._url_col   = tk.StringVar(value="AV")
        self._progress  = tk.StringVar(value="進度：-")
        self._url_var   = tk.StringVar(value="目前網址：-")
        self._status    = tk.StringVar(value="狀態：就緒")

        self._processor: ExcelProcessor | None = None
        self._build_ui()
        self._load_session()

    # ------------------------------------------------------------------
    def _build_ui(self):
        root = self._root
        pad = {"padx": 10, "pady": 4}

        # 檔案列
        f_file = tk.Frame(root)
        f_file.pack(fill="x", **pad)
        tk.Button(f_file, text="選擇 Excel 檔案", command=self._choose_file,
                  width=16).pack(side="left")
        tk.Label(f_file, textvariable=self._filepath, anchor="w",
                 wraplength=420).pack(side="left", fill="x", expand=True, padx=(8, 0))

        # 起始列 + 網址欄位
        f_row = tk.Frame(root)
        f_row.pack(fill="x", **pad)
        tk.Label(f_row, text="從第").pack(side="left")
        tk.Spinbox(f_row, textvariable=self._start_row,
                   from_=2, to=999999, width=7).pack(side="left", padx=(4, 4))
        tk.Label(f_row, text="列開始").pack(side="left")
        tk.Label(f_row, text="網址欄位：").pack(side="left", padx=(16, 0))
        tk.Entry(f_row, textvariable=self._url_col, width=5).pack(side="left")

        # 按鈕列
        f_btn = tk.Frame(root)
        f_btn.pack(fill="x", **pad)
        self._btn_start = tk.Button(f_btn, text="開始", width=10,
                                    command=self._start, state="disabled")
        self._btn_start.pack(side="left")
        self._btn_start_single = tk.Button(f_btn, text="只處理設定的開始列", width=18,
                                           command=self._start_single, state="disabled")
        self._btn_start_single.pack(side="left", padx=(8, 0))
        self._btn_stop = tk.Button(f_btn, text="停止", width=10,
                                   command=self._stop, state="disabled")
        self._btn_stop.pack(side="left", padx=(8, 0))

        # 資訊列
        for var in (self._progress, self._url_var, self._status):
            tk.Label(root, textvariable=var, anchor="w",
                     wraplength=560).pack(fill="x", **pad)

    # ------------------------------------------------------------------
    def _load_session(self):
        data = load_session()
        filepath = data.get("filepath", "")
        if filepath and __import__("os").path.isfile(filepath):
            self._filepath.set(filepath)
            self._start_row.set(data.get("start_row", 2))
            self._btn_start.config(state="normal")
            self._btn_start_single.config(state="normal")
        self._url_col.set(data.get("url_col", "AV"))

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xlsm"), ("所有檔案", "*.*")],
        )
        if path:
            if path != self._filepath.get():
                self._start_row.set(2)
            self._filepath.set(path)
            self._btn_start.config(state="normal")
            self._btn_start_single.config(state="normal")
            self._status.set("狀態：就緒")
            save_session(path, self._start_row.get(), self._url_col.get())

    def _start(self):
        self._run_processor(single_row=False)

    def _start_single(self):
        self._run_processor(single_row=True)

    def _run_processor(self, single_row: bool):
        filepath = self._filepath.get()
        if not filepath or filepath == "（尚未選擇檔案）":
            messagebox.showwarning("提示", "請先選擇 Excel 檔案")
            return

        self._btn_start.config(state="disabled")
        self._btn_start_single.config(state="disabled")
        self._btn_stop.config(state="normal")
        self._status.set("狀態：處理中...")
        self._progress.set("進度：-")

        self._processor = ExcelProcessor(
            filepath=filepath,
            progress_cb=self._on_progress,
            status_cb=self._on_status,
            start_row=self._start_row.get(),
            url_col=self._url_col.get(),
            single_row=single_row,
        )
        self._processor.start()

    def _stop(self):
        if self._processor:
            self._processor.stop()
        self._btn_stop.config(state="disabled")

    # ------------------------------------------------------------------
    def _on_progress(self, row_num: int, url: str):
        self._root.after(0, lambda: self._progress.set(f"進度：正在處理第 {row_num} 列"))
        self._root.after(0, lambda: self._url_var.set(f"目前網址：{url}"))

    def _on_status(self, msg: str):
        self._root.after(0, lambda: self._status.set(f"狀態：{msg}"))
        self._root.after(0, self._reset_buttons)

    def _reset_buttons(self):
        self._btn_start.config(state="normal")
        self._btn_start_single.config(state="normal")
        self._btn_stop.config(state="disabled")

    # ------------------------------------------------------------------
    def run(self):
        self._root.mainloop()
