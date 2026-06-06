import os
from playwright.sync_api import sync_playwright, Page, BrowserContext, Playwright

_PROFILE_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "chrome_profile")
)


class BrowserSession:
    def __init__(self):
        self._pw: Playwright | None = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None

    def __enter__(self) -> "BrowserSession":
        self._pw = sync_playwright().start()
        self._context = self._pw.chromium.launch_persistent_context(
            user_data_dir=_PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=["--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        self._page = self._context.new_page()
        return self

    def __exit__(self, *_):
        try:
            self._context.close()
        finally:
            self._pw.stop()

    def fetch_text(self, url: str) -> str:
        self._page.goto(url, wait_until="domcontentloaded", timeout=30000)
        text = self._page.inner_text("body")
        if len(text) < 500 and "cloudflare" in text.lower():
            self._page.wait_for_function(
                "document.body.innerText.length > 500", timeout=120000
            )
            text = self._page.inner_text("body")
        return self._trim_to_title(text)

    def _trim_to_title(self, text: str) -> str:
        """Drop navigation/breadcrumb — start from the standalone patent title (h1)."""
        try:
            h1 = self._page.inner_text("h1").strip()
            if not h1:
                return text
            keyword = h1[:40]
            first = text.find(keyword)
            if first == -1:
                return text
            # Title appears first inside breadcrumb, then again as a standalone line
            second = text.find(keyword, first + len(keyword))
            text = text[second if second != -1 else first:]
        except Exception:
            pass
        return self._trim_after_classifications(text)

    def _trim_after_classifications(self, text: str) -> str:
        """Cut off footer — end after the last classification line."""
        for marker in ("International Classification:", "Current U.S. Class:"):
            idx = text.find(marker)
            if idx != -1:
                end = text.find("\n", idx)
                return text[:end].rstrip() if end != -1 else text
        # Fallback: cut before the website footer
        for footer in ("\n\nAsk a Lawyer", "\n\nFind a Lawyer"):
            idx = text.find(footer)
            if idx != -1:
                return text[:idx].rstrip()
        return text
