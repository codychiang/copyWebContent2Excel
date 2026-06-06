from __future__ import annotations

import datetime

# ─── 全域開關 ───────────────────────────────────────────────
DEBUG = True   # 改成 False 關閉所有 debug 輸出
# ────────────────────────────────────────────────────────────


def dlog(msg: str) -> None:
    if DEBUG:
        ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[DEBUG {ts}] {msg}", flush=True)
