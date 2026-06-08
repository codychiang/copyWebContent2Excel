import os
from playwright.sync_api import sync_playwright, Page, BrowserContext, Playwright

from .debug import dlog

class BrowserSession:
    def __init__(self, xlsx_filepath: str):
        base = os.path.splitext(os.path.abspath(xlsx_filepath))[0]
        self._profile_dir = base + "_chrome_profile"
        self._pw: Playwright | None = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None

    def __enter__(self) -> "BrowserSession":
        dlog(f"BrowserSession.__enter__: 啟動 Playwright，profile={self._profile_dir}")
        self._pw = sync_playwright().start()
        dlog("BrowserSession.__enter__: launch_persistent_context ...")
        self._context = self._pw.chromium.launch_persistent_context(
            user_data_dir=self._profile_dir,
            channel="chrome",
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        dlog("BrowserSession.__enter__: new_page ...")
        self._page = self._context.new_page()
        dlog("BrowserSession.__enter__: 完成")
        return self

    def __exit__(self, *_):
        dlog("BrowserSession.__exit__: 關閉 context ...")
        try:
            self._context.close()
        finally:
            self._pw.stop()
        dlog("BrowserSession.__exit__: 完成")

    def fetch_text(self, url: str) -> str:
        dlog(f"fetch_text: goto {url}")
        self._page.goto(url, wait_until="domcontentloaded", timeout=30000)
        dlog("fetch_text: page loaded，讀取 body text ...")
        text = self._page.inner_text("body")
        dlog(f"fetch_text: body 長度={len(text)}")
        if len(text) == 0 or (len(text) < 500 and "cloudflare" in text.lower()):
            dlog("fetch_text: 偵測到 Cloudflare，等待頁面完整載入 (最長 120s) ...")
            self._page.wait_for_function(
                "document.body.innerText.length > 500", timeout=120000
            )
            text = self._page.inner_text("body")
            dlog(f"fetch_text: Cloudflare 後 body 長度={len(text)}")
        result = self._trim_to_title(text)
        dlog(f"fetch_text: 修剪後長度={len(result)}")
        return result

    def _trim_to_title(self, text: str) -> str:
        """Drop navigation/breadcrumb — start from the standalone patent title (h1)."""
        try:
            h1 = self._page.inner_text("h1").strip()
            dlog(f"_trim_to_title: h1={repr(h1[:60]) if h1 else '(空)'}")
            if not h1:
                return text
            keyword = h1[:40]
            first = text.find(keyword)
            if first == -1:
                dlog("_trim_to_title: 找不到 h1 關鍵字，不修剪")
                return text
            second = text.find(keyword, first + len(keyword))
            dlog(f"_trim_to_title: first={first}, second={second}，從 {'second' if second != -1 else 'first'} 開始截取")
            text = text[second if second != -1 else first:]
        except Exception as e:
            dlog(f"_trim_to_title: 例外 {e}")
        return self._trim_after_classifications(text)

    def _trim_after_classifications(self, text: str) -> str:
        """Cut off footer — end after the last classification line."""
        for marker in ("International Classification:", "Current U.S. Class:"):
            idx = text.find(marker)
            if idx != -1:
                end = text.find("\n", idx)
                dlog(f"_trim_after_classifications: 找到 '{marker}' at {idx}，截斷 at {end}")
                return text[:end].rstrip() if end != -1 else text
        for footer in ("\n\nAsk a Lawyer", "\n\nFind a Lawyer"):
            idx = text.find(footer)
            if idx != -1:
                dlog(f"_trim_after_classifications: 找到 footer '{footer.strip()}' at {idx}，截斷")
                return text[:idx].rstrip()
        dlog("_trim_after_classifications: 無截斷標記，回傳原文")
        return text
