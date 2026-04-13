@echo off
title eToro Dashboard Refresh
setlocal

cd /d "%~dp0"

:: ── Find Python ───────────────────────────────────────────────────────────────
set PYTHON=
if exist "D:\Anaconda\python.exe"         set PYTHON=D:\Anaconda\python.exe
if exist "C:\Anaconda3\python.exe"        set PYTHON=C:\Anaconda3\python.exe
if exist "C:\ProgramData\Anaconda3\python.exe" set PYTHON=C:\ProgramData\Anaconda3\python.exe
if "%PYTHON%"=="" (
    where python >nul 2>&1
    if not errorlevel 1 (set PYTHON=python) else (
        echo ERROR: Python not found. & exit /b 1
    )
)

:: ── Load .env ─────────────────────────────────────────────────────────────────
if exist etoro.env (
    for /f "usebackq eol=# tokens=1,* delims==" %%A in ("etoro.env") do (
        if not "%%A"=="" if not "%%A"==" " set "%%A=%%B"
    )
)

:: ── Run dashboard only ────────────────────────────────────────────────────────
"%PYTHON%" run_all.py --dash >> logs\dashboard_refresh.log 2>&1

:: ── Copy dashboard to Google Drive Upload folder ─────────────────────────────
robocopy "C:\Users\Neil\ClaudeCode\eToro" "C:\Users\Neil\My Drive\Upload" *.html /njh /njs /nfl /ndl >> logs\dashboard_refresh.log 2>&1

endlocal
