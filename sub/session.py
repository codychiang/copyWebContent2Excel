from __future__ import annotations

import configparser
import os

_INI_FILE = os.path.join(os.path.dirname(__file__), "..", "setting.ini")
_SECTION  = "session"


def load_session() -> dict:
    cfg = configparser.ConfigParser()
    cfg.read(_INI_FILE, encoding="utf-8-sig")
    if not cfg.has_section(_SECTION):
        return {}
    return {
        "filepath":  cfg.get(_SECTION, "filepath",  fallback=""),
        "start_row": cfg.getint(_SECTION, "start_row", fallback=2),
        "url_col":   cfg.get(_SECTION, "url_col",   fallback="AV"),
    }


def save_session(filepath: str, start_row: int, url_col: str = "AV") -> None:
    cfg = configparser.ConfigParser()
    cfg.read(_INI_FILE, encoding="utf-8-sig")
    if not cfg.has_section(_SECTION):
        cfg.add_section(_SECTION)
    cfg.set(_SECTION, "filepath",  filepath)
    cfg.set(_SECTION, "start_row", str(start_row))
    cfg.set(_SECTION, "url_col",   url_col.upper().strip())
    with open(_INI_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)
