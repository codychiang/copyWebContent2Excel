@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

python make_splash.py
pyinstaller --onedir --windowed --noconfirm --splash splash.png --name copyWebContent2Excel --add-data "windows.code-workspace;." main.py
pause
