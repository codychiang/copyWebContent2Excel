"""Read windows.code-workspace and generate defines.nsh for NSIS."""
import json
import os
import re
import sys

_workspace = os.path.join(os.path.dirname(__file__), "..", "windows.code-workspace")

try:
    text = open(_workspace, encoding="utf-8").read()
    text = re.sub(r"//[^\n]*", "", text)          # 去除 // 行注釋
    text = re.sub(r",\s*([}\]])", r"\1", text)    # 去除尾逗號（JSONC → JSON）
    settings = json.loads(text).get("settings", {})
    name    = settings.get("app.name",    "")
    version = settings.get("app.version", "?")
    author  = settings.get("app.author",  "")
except Exception as e:
    print(f"錯誤：{e}", file=sys.stderr)
    sys.exit(1)

out = os.path.join(os.path.dirname(__file__), "defines.nsh")
with open(out, "w", encoding="utf-8") as f:
    f.write(f'!define APP_NAME    "{name}"\n')
    f.write(f'!define APP_VERSION "{version}"\n')
    f.write(f'!define APP_AUTHOR  "{author}"\n')

print(f"defines.nsh OK: {name} {version} / {author}")
