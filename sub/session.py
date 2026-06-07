from __future__ import annotations

import configparser
import json
import os
import re
import sys

_INI_FILE = os.path.join(os.path.dirname(__file__), "..", "setting.ini")

# 打包後檔案在 sys._MEIPASS，開發時在專案根目錄
_BASE = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.join(os.path.dirname(__file__), "..")
_WORKSPACE_FILE = os.path.join(_BASE, "windows.code-workspace")
_SECTION        = "session"

_DEFAULT_ALARM = r"C:\Windows\Media\Alarm05.wav"


def load_session() -> dict:
    cfg = configparser.ConfigParser()
    cfg.read(_INI_FILE, encoding="utf-8-sig")
    if not cfg.has_section(_SECTION):
        return {}
    return {
        "filepath":  cfg.get(_SECTION, "filepath",  fallback=""),
        "start_row": cfg.getint(_SECTION, "start_row", fallback=2),
        "url_col":   cfg.get(_SECTION, "url_col",   fallback="AV"),
        "alarm_wav": cfg.get(_SECTION, "alarm_wav", fallback=_DEFAULT_ALARM),
    }


def _jsonc_to_json(text: str) -> str:
    text = re.sub(r"//[^\n]*", "", text)         # 去除 // 行注釋
    text = re.sub(r",\s*([}\]])", r"\1", text)   # 去除尾逗號（JSONC → JSON）
    return text


def load_app_settings() -> dict:
    """Return app.name and app.version from windows.code-workspace settings."""
    try:
        text = open(_WORKSPACE_FILE, encoding="utf-8").read()
        settings = json.loads(_jsonc_to_json(text)).get("settings", {})
        return {
            "name":    settings.get("app.name",    ""),
            "version": settings.get("app.version", "?"),
            "author":  settings.get("app.author",  ""),
        }
    except Exception:
        return {"name": "copyWebContent2Excel", "version": "?", "author": ""}


def load_version() -> str:
    return load_app_settings()["version"]


def save_session(filepath: str, start_row: int, url_col: str = "AV",
                 alarm_wav: str = _DEFAULT_ALARM) -> None:
    cfg = configparser.ConfigParser()
    cfg.read(_INI_FILE, encoding="utf-8-sig")
    if not cfg.has_section(_SECTION):
        cfg.add_section(_SECTION)
    cfg.set(_SECTION, "filepath",  filepath)
    cfg.set(_SECTION, "start_row", str(start_row))
    cfg.set(_SECTION, "url_col",   url_col.upper().strip())
    cfg.set(_SECTION, "alarm_wav", alarm_wav)
    with open(_INI_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)
