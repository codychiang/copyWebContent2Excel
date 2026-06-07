from __future__ import annotations

import tkinter as tk
import winsound
from tkinter import filedialog, messagebox, font as tkfont, ttk

from .processor import ExcelProcessor, TxtFolderFiller
from .session import load_app_settings, load_session, save_session

# ── Palette ───────────────────────────────────────────────────────
BG       = "#f0f2f5"   # 頁面背景（淡灰）
CARD     = "#ffffff"   # 卡片白
BORDER   = "#dde1e7"   # 邊框
TEXT     = "#1a1d23"   # 主要文字
MUTED    = "#6b7280"   # 次要文字
BLUE     = "#2563eb"   # 主色（按鈕 / 資訊）
GREEN    = "#16a34a"   # 成功
RED      = "#dc2626"   # 危險 / 錯誤
ORANGE   = "#ea580c"   # 警告 / 進行中
BLUE_LT  = "#eff6ff"   # 藍色淺底
GREEN_LT = "#f0fdf4"   # 綠色淺底
RED_LT   = "#fef2f2"   # 紅色淺底

FONT     = ("Segoe UI", 10)
FONT_SM  = ("Segoe UI", 9)
FONT_LG  = ("Segoe UI", 12, "bold")
FONT_MONO = ("Consolas", 9)


def _btn(parent, text, cmd, bg, fg="#ffffff", width=None, state="normal"):
    """3-D raised button: RAISED at rest → SUNKEN on press."""
    import math

    def _lighten(hex_col, factor=1.25):
        h = hex_col.lstrip("#")
        r, g, b = (int(h[i:i+2], 16) for i in (0, 2, 4))
        return "#{:02x}{:02x}{:02x}".format(
            min(255, math.floor(r*factor)),
            min(255, math.floor(g*factor)),
            min(255, math.floor(b*factor)))

    def _darken(hex_col, factor=0.78):
        h = hex_col.lstrip("#")
        r, g, b = (int(h[i:i+2], 16) for i in (0, 2, 4))
        return "#{:02x}{:02x}{:02x}".format(
            math.floor(r*factor), math.floor(g*factor), math.floor(b*factor))

    kw = dict(
        text=text, command=cmd,
        bg=bg, fg=fg,
        activebackground=_darken(bg, 0.85),
        activeforeground=fg,
        disabledforeground="#a8b0bb",
        font=("Segoe UI", 10, "bold"),
        relief=tk.RAISED, bd=3,
        highlightbackground=_darken(bg, 0.65),
        highlightcolor=_lighten(bg, 1.35),
        highlightthickness=1,
        padx=16, pady=6,
        cursor="hand2",
        overrelief=tk.RAISED,
    )
    if width:
        kw["width"] = width
    b = tk.Button(parent, **kw)
    b.config(state=state)
    return b


class App:
    def __init__(self):
        self._root = tk.Tk()
        self._root.title("Patent Content → Excel")
        self._root.configure(bg=BG)
        self._root.resizable(True, True)
        self._root.minsize(740, 580)

        # ttk 捲軸樣式
        sty = ttk.Style(self._root)
        sty.theme_use("clam")
        sty.configure("Light.Vertical.TScrollbar",
                       background=BORDER, troughcolor=BG,
                       arrowcolor=MUTED, bordercolor=BORDER,
                       darkcolor=BORDER, lightcolor=CARD)

        self._filepath  = tk.StringVar(value="尚未選擇檔案")
        self._start_row = tk.IntVar(value=2)
        self._url_col   = tk.StringVar(value="AV")
        self._progress  = tk.StringVar(value="—")
        self._url_var   = tk.StringVar(value="—")
        self._status    = tk.StringVar(value="就緒")
        self._status_lbl: tk.Label | None = None

        self._processor: ExcelProcessor | None = None
        self._filler: TxtFolderFiller | None = None
        self._alert_playing = False
        self._is_crawling = False
        self._is_filling  = False
        self._alarm_wav = r"C:\Windows\Media\Alarm05.wav"

        self._build_ui()
        self._build_menu()
        self._load_session()

    # ── Build ─────────────────────────────────────────────────────
    def _build_ui(self):
        root = self._root

        # ────────────────────────────────────────────────
        # 頁首 (白色卡片)
        # ────────────────────────────────────────────────
        header = tk.Frame(root, bg=CARD, pady=14)
        header.pack(fill="x")
        tk.Frame(header, bg=BLUE, width=4).pack(side="left", fill="y")
        tk.Label(header, text="Patent Content → Excel",
                 font=FONT_LG, bg=CARD, fg=TEXT,
                 padx=16).pack(side="left")

        tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

        # ────────────────────────────────────────────────
        # 檔案選擇列
        # ────────────────────────────────────────────────
        file_row = tk.Frame(root, bg=BG, pady=14)
        file_row.pack(fill="x", padx=20)

        self._btn_choose = _btn(file_row, "選擇 Excel 檔案",
                                self._choose_file, bg=BLUE, width=14)
        self._btn_choose.pack(side="left")

        self._lbl_file = tk.Label(
            file_row, textvariable=self._filepath,
            font=FONT_SM, bg=BG, fg=MUTED,
            anchor="w", wraplength=420, justify="left")
        self._lbl_file.pack(side="left", padx=(12, 0))

        # ────────────────────────────────────────────────
        # 設定列 (白色卡片)
        # ────────────────────────────────────────────────
        tk.Frame(root, bg=BORDER, height=1).pack(fill="x")
        cfg_card = tk.Frame(root, bg=CARD, pady=12)
        cfg_card.pack(fill="x")
        tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

        cfg = tk.Frame(cfg_card, bg=CARD)
        cfg.pack(padx=20, anchor="w")

        tk.Label(cfg, text="起始列", font=FONT_SM, bg=CARD, fg=MUTED).grid(
            row=0, column=0, sticky="w")
        self._spin = tk.Spinbox(
            cfg, textvariable=self._start_row,
            from_=1, to=999999, width=8, font=FONT,
            bg=CARD, fg=TEXT, buttonbackground=BG,
            relief=tk.SOLID, bd=1, highlightthickness=0,
            insertbackground=TEXT)
        self._spin.grid(row=0, column=1, padx=(8, 24), sticky="w")

        tk.Label(cfg, text="網址欄位", font=FONT_SM, bg=CARD, fg=MUTED).grid(
            row=0, column=2, sticky="w")
        self._entry_col = tk.Entry(
            cfg, textvariable=self._url_col, width=6, font=FONT,
            bg=CARD, fg=TEXT, relief=tk.SOLID, bd=1, highlightthickness=0,
            insertbackground=TEXT)
        self._entry_col.grid(row=0, column=3, padx=(8, 0), sticky="w")

        # ────────────────────────────────────────────────
        # 操作按鈕
        # ────────────────────────────────────────────────
        btn_area = tk.Frame(root, bg=BG, pady=16)
        btn_area.pack(fill="x", padx=20)

        # 列 1：抓取
        row1 = tk.Frame(btn_area, bg=BG)
        row1.pack(fill="x")

        self._btn_start = _btn(
            row1, "開始抓取", self._start, bg=GREEN, width=10, state="disabled")
        self._btn_start.pack(side="left")

        self._btn_start_single = _btn(
            row1, "只抓取此列", self._start_single,
            bg="#e2e8f0", fg=TEXT, width=10, state="disabled")
        self._btn_start_single.pack(side="left", padx=(10, 0))

        self._btn_stop = _btn(
            row1, "停止", self._stop, bg=RED, width=7, state="disabled")
        self._btn_stop.pack(side="left", padx=(10, 0))

        # 列 2：Excel 寫入
        row2 = tk.Frame(btn_area, bg=BG)
        row2.pack(fill="x", pady=(10, 0))

        self._btn_fill = _btn(
            row2, "填入 Excel", self._fill_excel, bg=BLUE, width=10, state="disabled")
        self._btn_fill.pack(side="left")

        self._btn_stop_fill = _btn(
            row2, "停止填入", self._stop_fill, bg=RED, width=8, state="disabled")
        self._btn_stop_fill.pack(side="left", padx=(10, 0))

        tk.Label(row2, text="掃描 txt 資料夾並寫入 xlsx",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(12, 0))

        self._btn_preview = _btn(
            row2, "試聽警報", self._preview_alert, bg=ORANGE, width=8)
        self._btn_preview.pack(side="right")

        # ────────────────────────────────────────────────
        # 狀態列 (白色卡片)
        # ────────────────────────────────────────────────
        tk.Frame(root, bg=BORDER, height=1).pack(fill="x")
        info_card = tk.Frame(root, bg=CARD, pady=12)
        info_card.pack(fill="x")
        tk.Frame(root, bg=BORDER, height=1).pack(fill="x")

        info = tk.Frame(info_card, bg=CARD)
        info.pack(fill="x", padx=20)

        def info_row(label, var, default_fg):
            row = tk.Frame(info, bg=CARD)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, width=5, anchor="w",
                     font=FONT_SM, bg=CARD, fg=MUTED).pack(side="left")
            lbl = tk.Label(row, textvariable=var, anchor="w",
                           font=FONT_SM, bg=CARD, fg=default_fg,
                           wraplength=560, justify="left")
            lbl.pack(side="left", fill="x")
            return lbl

        info_row("進度", self._progress, BLUE)
        info_row("網址", self._url_var, MUTED)
        self._status_lbl = info_row("狀態", self._status, GREEN)

        # ────────────────────────────────────────────────
        # 內容預覽
        # ────────────────────────────────────────────────
        preview_wrap = tk.Frame(root, bg=BG)
        preview_wrap.pack(fill="both", expand=True, padx=20, pady=12)

        tk.Label(preview_wrap, text="內容預覽",
                 font=FONT_SM, bg=BG, fg=MUTED).pack(anchor="w", pady=(0, 6))

        txt_border = tk.Frame(preview_wrap, bg=BORDER, bd=0)
        txt_border.pack(fill="both", expand=True)

        inner = tk.Frame(txt_border, bg=CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        scrollbar = ttk.Scrollbar(inner, orient="vertical",
                                  style="Light.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")

        self._text = tk.Text(
            inner, wrap="word", state="disabled",
            font=FONT_MONO,
            bg=CARD, fg=TEXT,
            insertbackground=TEXT,
            selectbackground="#bfdbfe",
            selectforeground=TEXT,
            relief=tk.FLAT, bd=0,
            padx=12, pady=10,
            yscrollcommand=scrollbar.set)
        self._text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self._text.yview)

    # ── Menu ──────────────────────────────────────────────────────
    def _build_menu(self):
        menubar = tk.Menu(self._root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Exit", command=self._root.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        util_menu = tk.Menu(menubar, tearoff=0)
        util_menu.add_command(label="選擇 Alarm 音檔", command=self._choose_alarm_file)
        self._topmost_var = tk.BooleanVar(value=False)
        util_menu.add_checkbutton(label="顯示在上層", variable=self._topmost_var,
                                  command=self._toggle_topmost)
        menubar.add_cascade(label="Utility", menu=util_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self._root.config(menu=menubar)

    def _choose_alarm_file(self):
        import os
        path = filedialog.askopenfilename(
            title="選擇警報音檔",
            initialdir=os.path.dirname(self._alarm_wav),
            initialfile=os.path.basename(self._alarm_wav),
            filetypes=[("WAV 音檔", "*.wav"), ("所有檔案", "*.*")],
        )
        if path:
            self._alarm_wav = path
            save_session(
                self._filepath.get() if self._filepath.get() != "尚未選擇檔案" else "",
                self._start_row.get(),
                self._url_col.get(),
                alarm_wav=path,
            )

    def _toggle_topmost(self):
        self._root.attributes("-topmost", self._topmost_var.get())

    def _show_about(self):
        info = load_app_settings()
        messagebox.showinfo(
            "About",
            f"{info['name']}\n版本：{info['version']}\n作者：{info['author']}\n\n將專利頁面內容擷取並寫入 Excel。",
        )

    # ── Session ───────────────────────────────────────────────────
    def _load_session(self):
        data = load_session()
        filepath = data.get("filepath", "")
        if filepath and __import__("os").path.isfile(filepath):
            self._filepath.set(filepath)
            self._start_row.set(data.get("start_row", 2))
            self._btn_start.config(state="normal")
            self._btn_start_single.config(state="normal")
            self._btn_fill.config(state="normal")
        self._url_col.set(data.get("url_col", "AV"))
        self._alarm_wav = data.get("alarm_wav", self._alarm_wav)

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
            self._btn_fill.config(state="normal")
            self._set_status("就緒", GREEN)

    # ── Actions ───────────────────────────────────────────────────
    def _start(self):
        self._run_processor(single_row=False)

    def _start_single(self):
        self._run_processor(single_row=True)

    def _run_processor(self, single_row: bool):
        filepath = self._filepath.get()
        if not filepath or filepath == "尚未選擇檔案":
            messagebox.showwarning("提示", "請先選擇 Excel 檔案")
            return

        self._stop_alert()
        save_session(filepath, self._start_row.get(), self._url_col.get(),
                     alarm_wav=self._alarm_wav)
        self._is_crawling = True
        self._set_busy()
        self._set_status("處理中…", ORANGE)
        self._progress.set("—")

        self._processor = ExcelProcessor(
            filepath=filepath,
            progress_cb=self._on_progress,
            status_cb=self._on_status,
            content_cb=self._on_content,
            error_cb=self._on_fetch_error,
            start_row=self._start_row.get(),
            url_col=self._url_col.get(),
            single_row=single_row,
        )
        self._processor.start()

    def _stop(self):
        self._stop_alert()
        if self._processor:
            self._processor.stop()
        self._btn_stop.config(state="disabled")

    def _fill_excel(self):
        filepath = self._filepath.get()
        if not filepath or filepath == "尚未選擇檔案":
            messagebox.showwarning("提示", "請先選擇 Excel 檔案")
            return

        self._stop_alert()
        self._is_filling = True
        self._set_busy()
        self._set_status("掃描 txt 資料夾…", ORANGE)

        self._filler = TxtFolderFiller(
            filepath=filepath,
            url_col=self._url_col.get(),
            status_cb=self._on_status,
        )
        self._filler.start()

    def _stop_fill(self):
        if self._filler:
            self._filler.stop()

    # ── Busy / reset ─────────────────────────────────────────────
    def _set_busy(self):
        self._btn_start.config(state="disabled")
        self._btn_start_single.config(state="disabled")
        self._btn_fill.config(state="disabled")
        self._btn_stop.config(state="normal" if self._is_crawling else "disabled")
        self._btn_stop_fill.config(state="normal" if self._is_filling else "disabled")

    def _reset_buttons(self):
        has_file = self._filepath.get() not in ("", "尚未選擇檔案")
        s = "normal" if has_file else "disabled"
        self._btn_start.config(state=s)
        self._btn_start_single.config(state=s)
        self._btn_fill.config(state=s)
        self._btn_stop.config(state="disabled")
        self._btn_stop_fill.config(state="disabled")

    def _set_status(self, msg: str, color: str = TEXT):
        self._status.set(msg)
        if self._status_lbl:
            self._status_lbl.config(fg=color)

    # ── Alert ─────────────────────────────────────────────────────
    def _preview_alert(self):
        winsound.PlaySound(self._alarm_wav, winsound.SND_FILENAME | winsound.SND_ASYNC)

    def _play_alert(self):
        self._alert_playing = True
        winsound.PlaySound(
            self._alarm_wav, winsound.SND_FILENAME | winsound.SND_LOOP | winsound.SND_ASYNC)
        self._btn_preview.config(text="停止警報", command=self._stop_alert)

    def _stop_alert(self):
        if self._alert_playing:
            winsound.PlaySound(None, winsound.SND_PURGE)
            self._alert_playing = False
        self._btn_preview.config(text="試聽警報", command=self._preview_alert)

    # ── Callbacks ─────────────────────────────────────────────────
    def _on_fetch_error(self, row_num: int, url: str, msg: str):
        def _update():
            self._progress.set(f"第 {row_num} 列擷取失敗（已停止）")
            self._url_var.set(url)
            self._set_status(f"擷取錯誤 — {msg}", RED)
            self._is_crawling = False
            self._is_filling  = False
            self._reset_buttons()
            self._play_alert()
        self._root.after(0, _update)

    def _on_content(self, text: str):
        def _update():
            self._text.config(state="normal")
            self._text.delete("1.0", "end")
            self._text.insert("1.0", text)
            self._text.config(state="disabled")
            self._text.see("1.0")
        self._root.after(0, _update)

    def _on_progress(self, row_num: int, url: str, current: int, total: int):
        self._root.after(
            0, lambda r=row_num, t=total: self._progress.set(f"正在處理 {r} 列 / {t}"))
        self._root.after(0, lambda: self._url_var.set(url))

    _TERMINAL = ("完成", "已停止", "錯誤", "找不到", "此為標題列", "沒有需要")

    def _on_status(self, msg: str):
        color = GREEN if "完成" in msg else ORANGE if ("寫入" in msg or "掃描" in msg) else TEXT
        is_done = any(t in msg for t in self._TERMINAL)

        def _update():
            self._set_status(msg, color)
            if is_done:
                self._is_crawling = False
                self._is_filling  = False
                self._reset_buttons()

        self._root.after(0, _update)

    # ── Run ───────────────────────────────────────────────────────
    def run(self):
        self._root.attributes("-topmost", True)
        self._root.lift()
        self._root.focus_force()
        self._root.after(300, lambda: self._root.attributes("-topmost", False))
        self._root.mainloop()
