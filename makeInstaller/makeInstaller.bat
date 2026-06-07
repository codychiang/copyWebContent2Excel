@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

:: 從 windows.code-workspace 讀取 python.defaultInterpreterPath
for /f "delims=" %%i in ('powershell -NoProfile -Command "$t=Get-Content '..\windows.code-workspace' -Raw; $t=$t -replace '//[^\n]*',''; $t=$t -replace ',\s*([}\]])','$1'; $p=($t|ConvertFrom-Json).settings.'python.defaultInterpreterPath'; $p=$p -replace '\$\{workspaceFolder\}',(Resolve-Path '..').Path; $p"') do set "PYTHON=%%i"

if "%PYTHON%"=="" (
    echo 無法從 workspace 取得 Python 路徑
    pause
    exit /b 1
)
echo Python: %PYTHON%

echo [1/3] 產生 defines.nsh...
set PYTHONUTF8=1
"%PYTHON%" gen_defines.py 2>&1
if errorlevel 1 (
    echo 錯誤：無法產生 defines.nsh
    pause
    exit /b 1
)

echo [2/3] 尋找 NSIS...
set "MAKENSIS="
if exist "C:\Program Files (x86)\NSIS\makensis.exe" set "MAKENSIS=C:\Program Files (x86)\NSIS\makensis.exe"
if exist "C:\Program Files\NSIS\makensis.exe"       set "MAKENSIS=C:\Program Files\NSIS\makensis.exe"

if "%MAKENSIS%"=="" (
    echo NSIS 未安裝，嘗試透過 winget 安裝...
    winget install NSIS.NSIS --silent
    if exist "C:\Program Files (x86)\NSIS\makensis.exe" set "MAKENSIS=C:\Program Files (x86)\NSIS\makensis.exe"
    if exist "C:\Program Files\NSIS\makensis.exe"       set "MAKENSIS=C:\Program Files\NSIS\makensis.exe"
)

if "%MAKENSIS%"=="" (
    echo 找不到 makensis.exe，請手動安裝 NSIS: https://nsis.sourceforge.io/
    pause
    exit /b 1
)

echo 使用: %MAKENSIS%
echo [3/3] 執行 NSIS...
if not exist "output" mkdir output
"%MAKENSIS%" app.nsi
pause
