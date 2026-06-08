try:
    import pyi_splash as _splash
except ImportError:
    _splash = None

from sub.fingerprint import verify
verify()

from sub.ui import App

if __name__ == "__main__":
    app = App()
    if _splash:
        _splash.close()
    app.run()
